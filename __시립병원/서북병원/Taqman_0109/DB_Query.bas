Attribute VB_Name = "DB_Query"
Option Explicit


Function GetUser(UserID As String, UserPass As String) As Boolean
    Dim i As Integer
    
    GetUser = False
    SQL = "SELECT UPASSWD "
    SQL = SQL & vbCrLf & "FROM USERMASTER"
    SQL = SQL & vbCrLf & "WHERE UID_1 = '" & Trim(UserID) & "' "
'            gstrQuy = "SELECT USER_ID, USER_NM, PWD "
'            gstrQuy = gstrQuy & vbCrLf & "  FROM TZUSERMSTN " '/HIS 사용자마스터 테이블(서북병원)
'            gstrQuy = gstrQuy & vbCrLf & " WHERE USER_ID = '" & Trim(txtUserID) & "' "
    res = db_select_Row(gServer, SQL)
    'strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            If gReadBuf(i) = UserPass Then
                GetUser = True
                Exit For
            End If
        Else
            Exit For
        End If
    Next

End Function

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
                    SQL = SQL & vbCrLf & "      RSLT_RPTR_ID = '" & gEquipCode & "_INF', "                                                           'Delta 체크"
                    SQL = SQL & vbCrLf & "      AMEN_ID = '" & gEquipCode & "_INF', "                                                           'Panic 체크"
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
                    SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(GetText(.vasTemp, iRow, 9)) & "' "
                    res = db_select_Col(gServer, SQL)
                
                    If gReadBuf(0) = "" Then
                        sCnt = "1"
                    Else
                        sCnt = CLng(gReadBuf(0)) + 1
                    End If
                    
                    If Trim(GetText(.vasTemp, iRow, 2)) <> "" Then
                        SQL = ""
                        SQL = SQL & vbCrLf & "INSERT INTO SPSLHQRST(EXMN_DY   ,EQPM_CD ,SBSN_CD ,LVL_CD  "
                        SQL = SQL & vbCrLf & "                     ,RSLT_SQNO ,EXMN_CD ,RSLT_DT ,RSLT_RPTR_ID "
                        SQL = SQL & vbCrLf & "                     ,RSLT_VALU ,SPCM_NO ,DEL_YN "
                        SQL = SQL & vbCrLf & "                     ,REGI_ID   ,RGST_DT ,AMEN_ID ,UPDT_DT) "
                        SQL = SQL & vbCrLf & "               VALUES('" & Trim(GetText(.vasTemp, iRow, 9)) & "', '" & Mid(lsID, 3, 3) & "', '" & Mid(lsID, 6, 3) & "', '" & Mid(lsID, 9, 1) & "', "
                        SQL = SQL & vbCrLf & "                      " & sCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gEquipCode & "_INF', "
                        'SQL = SQL & vbCrLf & "                      " & sCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gEquipCode & "_INF', "
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

''''-- 해당 환자 검사의 H/L, Delta, Panic 판정하기
'''Function GetDecision(ByVal argSpcRow As Integer, ByVal strBarno As String, ByVal iRow As Integer) As String
'''    Dim rs_Delta        As ADODB.Recordset
'''    Dim rs_DPRef        As ADODB.Recordset
'''    Dim strBefoRslt     As String
'''    Dim strDestRslt     As String
'''    Dim strHLVal        As String
'''    Dim strDelta        As String
'''    Dim strPanic        As String
'''    Dim strSex          As String
'''    Dim strHVal         As String
'''    Dim strLVal         As String
'''
'''    '-- 환자의 성별
'''    strSex = Trim(GetText(frmInterface.vasID, argSpcRow, colSex))
'''
'''    '-- 해당 환자의 참고치,델타,패닉 찾아오기
'''    SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW                 "
'''    SQL = SQL & vbCrLf & " FROM SPSLMFBIF                                                                                                                      "
'''    SQL = SQL & vbCrLf & " WHERE USE_STR_DY <= SYSDATE                                                                                                         "
'''    SQL = SQL & vbCrLf & "   AND USE_END_DY >= SYSDATE                                                                                                         "
'''    SQL = SQL & vbCrLf & "   and EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "'"
'''    Set rs_DPRef = cn_Ser.Execute(SQL)
'''    Do Until rs_DPRef.EOF
'''        '** 이전결과 조회 시작
'''        '-- 델타값을 계산하기 위한 이전결과 조회 (한달이내 결과값중 최근값만 조회한다.)
'''        SQL = ""
'''        SQL = SQL & vbCrLf & "SELECT B.SPCM_NO           BEFO_BCNO                                                               "
'''        SQL = SQL & vbCrLf & "     , B.EXMN_CD           BEFO_EXMN_CD                                                            "
'''        SQL = SQL & vbCrLf & "     , B.REAL_RSLT         BEFO_REAL_RSLT                                                          "
'''        SQL = SQL & vbCrLf & "     , B.VIEW_RSLT         BEFO_VIEW_RSLT                                                          "
'''        SQL = SQL & vbCrLf & "     , B.LAST_RPTG_DT     BEFO_FINL_DT                                                             "
'''        SQL = SQL & vbCrLf & "     , (SYSDATE - B.LAST_RPTG_DT)  DELTA_TERM_DT                                                   "  '오늘부터의 이전결과 기간을 구한다.
'''        SQL = SQL & vbCrLf & "     , B.PID               PID                                                                     "
'''        SQL = SQL & vbCrLf & "  FROM (SELECT MAX(B.LAST_RPTG_DT) LAST_RPTG_DT                                                    "
'''        SQL = SQL & vbCrLf & "             , B.EXMN_CD                                                                           "
'''        SQL = SQL & vbCrLf & "             , B.PID                                                                               "
'''        SQL = SQL & vbCrLf & "          FROM SPSLHRRST A, SPSLHRRST B                                                            "
'''        SQL = SQL & vbCrLf & "         WHERE A.SPCM_NO   <> B.SPCM_NO                                                            "
'''        SQL = SQL & vbCrLf & "           AND A.PID        = B.PID                                                                "
'''        SQL = SQL & vbCrLf & "           AND A.EXMN_CD    = B.EXMN_CD                                                            "
'''        SQL = SQL & vbCrLf & "           AND A.RCPN_DT   >= B.RCPN_DT                                                            "
'''        SQL = SQL & vbCrLf & "           AND B.LAST_RPTG_DT IS NOT NULL                                                          "
'''        SQL = SQL & vbCrLf & "           AND A.RSLT_STAT < '3'                                                                   "
'''        SQL = SQL & vbCrLf & "           AND A.SPCM_NO = FN_LABCVTBCNO('" & strBarno & "')                                       "
'''        SQL = SQL & vbCrLf & "         GROUP BY B.PID, B.EXMN_CD ) A, SPSLHRRST B                                                "
'''        SQL = SQL & vbCrLf & " WHERE A.PID = B.PID                                                                               "
'''        SQL = SQL & vbCrLf & "   AND A.LAST_RPTG_DT = B.LAST_RPTG_DT                                                             "
'''        SQL = SQL & vbCrLf & "   AND A.EXMN_CD = B.EXMN_CD                                                                       "
'''        SQL = SQL & vbCrLf & "   AND A.EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "         '검사코드"
'''        SQL = SQL & vbCrLf & "   AND B.LAST_RPTG_DT BETWEEN (SYSDATE-30) AND SYSDATE                "           '-- 30일 이내
'''        Set rs_Delta = cn_Ser.Execute(SQL)
'''        Do Until rs_Delta.EOF
'''            strBefoRslt = rs_Delta.Fields("BEFO_VIEW_RSLT")             '이전결과
'''            strDestRslt = Trim(GetText(frmInterface.vasTemp, iRow, 3))  '현재결과
'''
'''            '-- 성별로 판정결과 비교
'''            '-- 결과값이 수치일 경우에만 비교한다.
'''            If IsNumeric(strDestRslt) Then
'''                If strSex = "M" Then
'''                    If IsNumeric(rs_DPRef.Fields("MALE_HIGH")) Then
'''                        If CDbl(strDestRslt) > CDbl(rs_DPRef.Fields("MALE_HIGH")) Then
'''                            strHLVal = "H"
'''                        Else
'''                            strHLVal = ""
'''                        End If
'''                    Else
'''                        strHLVal = ""
'''                    End If
'''
'''                    If IsNumeric(rs_DPRef.Fields("MALE_LOW")) Then
'''                        If CDbl(strDestRslt) < CDbl(rs_DPRef.Fields("MALE_LOW")) Then
'''                            strHLVal = "L"
'''                        Else
'''                            strHLVal = ""
'''                        End If
'''                    Else
'''                        strHLVal = ""
'''                    End If
'''
'''                Else
'''                    If IsNumeric(rs_DPRef.Fields("FEML_HIGH")) Then
'''                        If CDbl(strDestRslt) > CDbl(rs_DPRef.Fields("FEML_HIGH")) Then
'''                            strHLVal = "H"
'''                        Else
'''                            strHLVal = ""
'''                        End If
'''                    Else
'''                        strHLVal = ""
'''                    End If
'''                    If IsNumeric(rs_DPRef.Fields("FEML_LOW")) Then
'''                        If CDbl(strDestRslt) < CDbl(rs_DPRef.Fields("FEML_LOW")) Then
'''                            strHLVal = "L"
'''                        Else
'''                            strHLVal = ""
'''                        End If
'''                    Else
'''                        strHLVal = ""
'''                    End If
'''                End If
'''            Else
'''                strHLVal = ""
'''            End If
'''
'''            '-- Delta 구분  (아래 로직이 맞는지 검증 필요함...必)
'''            '-- 결과값이 수치일 경우에만 비교한다.
'''            If IsNumeric(strDestRslt) Then
'''                Select Case Trim(rs_DPRef.Fields("DELT_DVSN"))
'''                    Case 0:     '0 사용안함
'''                            strDelta = ""
'''                    Case 1:     '1 변화차 = 현재결과 - 이전결과
'''                            strDelta = ""
'''                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'''                    Case 2:     '2 변화비율 = 변화차 / 이전결과 * 100
'''                            strDelta = ""
'''                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'''                            strDelta = (CDbl(strDelta) / CDbl(strBefoRslt)) * 100               '변화비율
'''                    Case 3:     '3 기간당 변화비율 = 변화비율 / 기간
'''                            strDelta = ""
'''                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'''                            strDelta = (CDbl(strDelta) / CDbl(strBefoRslt)) * 100               '변화비율
'''                            strDelta = strDelta / CInt(rs_Delta.Fields("DELTA_TERM_DT"))        '기간당 변화비율
'''                    Case 4:     '4 기간당 변화차 = 변화차 / 기간
'''                            strDelta = ""
'''                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'''                            strDelta = CDbl(strDelta) / CInt(rs_Delta.Fields("DELTA_TERM_DT"))  '기간당 변화차
'''                    Case 5:     '5 절대변화비율 = 변화차 / 이전결과
'''                            strDelta = ""
'''                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'''                            strDelta = CDbl(strDelta) / CDbl(strBefoRslt)                       '절대변화비율
'''                    Case Else:
'''                            strDelta = ""
'''                End Select
'''            Else
'''                strDelta = ""
'''            End If
'''
'''            '-- Panic 구분
'''            '-- 결과값이 수치일 경우에만 비교한다.
'''            If IsNumeric(strDestRslt) Then
'''                Select Case Trim(rs_DPRef.Fields("PANC_DVSN"))
'''                    Case 0:     '0 사용안함
'''                            strPanic = ""
'''                    Case 1:     '1 상한만
'''                            If IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
'''                                If CDbl(strDestRslt) > rs_DPRef.Fields("PANC_HIGH") Then
'''                                    strPanic = "P"
'''                                Else
'''                                    strPanic = ""
'''                                End If
'''                            Else
'''                                strPanic = ""
'''                            End If
'''                    Case 2:     '2 하한만
'''                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) Then
'''                                If CDbl(strDestRslt) < rs_DPRef.Fields("PANC_LOW") Then
'''                                    strPanic = "P"
'''                                Else
'''                                    strPanic = ""
'''                                End If
'''                            Else
'''                                strPanic = ""
'''                            End If
'''                    Case 3:     '3 모두 사용
'''                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) And IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
'''                                If (CDbl(strDestRslt) < rs_DPRef.Fields("PANC_LOW") Or CDbl(strDestRslt) > rs_DPRef.Fields("PANC_HIGH")) Then
'''                                    strPanic = "P"
'''                                Else
'''                                    strPanic = ""
'''                                End If
'''                            Else
'''                                strPanic = ""
'''                            End If
'''                    Case Else:
'''                            strPanic = ""
'''                End Select
'''            Else
'''                strPanic = ""
'''            End If
'''            rs_Delta.MoveNext
'''        Loop
'''
'''        rs_DPRef.MoveNext
'''    Loop
'''
'''    Set rs_DPRef = Nothing
'''
'''    GetDecision = strHLVal & "|" & strDelta & "|" & strPanic
'''
'''
'''End Function

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
    
    '-- 해당 환자의 참고치,델타,패닉 찾아오기
    SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW                 "
    SQL = SQL & vbCrLf & " FROM SPSLMFBIF                                                                                                                      "
    SQL = SQL & vbCrLf & " WHERE USE_STR_DY <= SYSDATE                                                                                                         "
    SQL = SQL & vbCrLf & "   AND USE_END_DY >= SYSDATE                                                                                                         "
    SQL = SQL & vbCrLf & "   and EXMN_CD = '" & Trim(strExamCode) & "' "
    Set rs_DPRef = cn_Ser.Execute(SQL)
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

'''Function Insert_Data(ByVal argSpcRow As Integer) As Integer
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
'''    Insert_Data = -1
'''
'''    lsID = ""
'''    lsID = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
'''    lsSpecNo = Trim(GetText(frmInterface.vasID, argSpcRow, colSpecNo))
'''    lsPid = Trim(GetText(frmInterface.vasID, argSpcRow, colPID))
'''    lsInsertTime = Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
'''
'''    If lsSpecNo = "" Then
'''        Exit Function
'''    End If
'''
'''    'Local에서 환자별로 결과값 가져오기
'''    ClearSpread frmInterface.vasTemp
'''
'''    SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
'''          " From pat_res " & vbCrLf & _
'''          " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''          " And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
'''          " And barcode = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
'''          " And diskno = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
'''          " And posno = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colPos)) & "' "
'''    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
'''
'''    If res = -1 Then
'''        SaveQuery SQL
'''        cn_Ser.RollbackTrans
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
'''        SQL = "SELECT RSLT_NO FROM SPSLHRRST "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '검사코드"
'''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
'''        res = db_select_Col(gServer, SQL)
'''        If res > 0 Then
'''            sCnt = CLng(gReadBuf(0)) + 1
'''
'''            '-- 결과값이 숫자값일 경우만 델타/패닉 판정을 한다.
'''            sResult = Trim(GetText(frmInterface.vasTemp, iRow, 3))
'''            If IsNumeric(sResult) Then
'''                Dim strDecision     As Variant
'''                Dim strBarcode      As String
'''
'''                strBarcode = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
''''                strDecision = GetDecision(argSpcRow, strBarcode, iRow)
'''                strDecision = GetDecision(argSpcRow, strBarcode, Trim(GetText(frmInterface.vasTemp, iRow, 2)), sResult)
'''                strDecision = Split(strDecision, "|")
'''            Else
'''                strDecision = "||"
'''                strDecision = Split(strDecision, "|")
'''            End If
'''
'''            SQL = "UPDATE SPSLHRRST "   '-- 결과테이블
'''            SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "                   '결과(장비결과)
'''            SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "                   '결과(수정결과)"
'''            SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & strDecision(0) & "', "                                                         'H/L 체크"
'''            SQL = SQL & vbCrLf & "       DLTA_YN = '" & strDecision(1) & "', "                                                           'Delta 체크"
'''            SQL = SQL & vbCrLf & "       PANC_YN = '" & strDecision(2) & "', "                                                           'Panic 체크"
'''            SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "', "                                                   '결과입력자"
'''            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
''''            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = 'test', "                                                   '중간보고자"
''''            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
''''            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = 'test', "                                                   '최종보고자"
''''            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "                                                        '결과수정자
'''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
'''            SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태"
'''            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "         '검사코드"
'''            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
'''            res = SendQuery(gServer, SQL)
'''
'''            If res < 0 Then
'''                SaveQuery SQL
'''                cn_Ser.RollbackTrans
'''                Exit Function
'''            End If
'''        End If
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
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1'"
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
'''    End If
'''
'''    SQL = ""
'''
'''    cn_Ser.CommitTrans
'''
'''    Insert_Data = 1
'''
'''    Exit Function
'''
'''Err:
'''    cn_Ser.RollbackTrans
'''
'''End Function


''''//////////////결과 저장 바꿈 (2011.10.11) - 효준
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
'''
'''    Dim Send_State      As String
'''    Dim SQL_LOCAL As String
'''
'''    With frmInterface
'''        'gComment_All = ""
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
'''
'''        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
'''              " From pat_res " & vbCrLf & _
'''              " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''              " And examdate = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
'''              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
'''              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
'''              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
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
''''        SQL = "SELECT EXMN_CD "
''''        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
''''        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
''''        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
''''        res = db_select_Col(gServer, SQL)
''''
''''        j = 0
''''        Do While gReadBuf(j) <> ""
''''            If ExamCode_Remark <> "" Then
''''                ExamCode_Remark = ExamCode_Remark & ",'" & gReadBuf(j) & "'"
''''            Else
''''                ExamCode_Remark = "'" & gReadBuf(j) & "'"
''''            End If
''''            j = j + 1
''''        Loop
''''        If ExamCode_Remark = "" Then ExamCode_Remark = "''"
'''
''''        For i = 1 To frmInterface.vasTemp.DataRowCnt
''''            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
''''        Next i
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
''''                gComment_Code = ""
'''
'''
'''                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
'''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
'''                res = db_select_Col(gServer, SQL)
'''
'''                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
'''
'''                sCnt = CLng(gReadBuf(0)) + 1
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
'''                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
''''                Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
'''                '/-----------------------------
'''
'''''                               SQL = "UPDATE SPSLHRRST "
'''''                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
'''''                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
'''''                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
'''''                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
'''''                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
'''''                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''''                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
'''''                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''''                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
'''
'''                SQL = "UPDATE SPSLHRRST "   '-- 결과테이블
'''                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "      '결과(장비결과)
'''                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "      '결과(수정결과)"
'''                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & strDecision(0) & "', "                                    'H/L 체크"
'''                SQL = SQL & vbCrLf & "       DLTA_YN = '" & strDecision(1) & "', "                                      'Delta 체크"
'''                SQL = SQL & vbCrLf & "       PANC_YN = '" & strDecision(2) & "', "                                      'Panic 체크"
'''                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
'''                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
'''
'''                '/////////// 혈액학 장비만 사용 ( 다른장비들은 결과 입력상태(= 1)로 함)
''''                If Mid(Trim(GetText(.vasTemp, iRow, 2)), 1, 2) = "L8" Then
''''                    Send_State = "1"
''''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
''''                Else
''''                    SQL_LOCAL = ""
''''                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "SELECT COUNT(EXAMCODE) FROM PAT_RES "
''''                    SQL_LOCAL = SQL_LOCAL & vbCrLf & " WHERE (REFFLAG <> '' OR PANICFLAG <> '' OR  DELTAFLAG <> '' ) "
''''                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND panicflag = 'P' "
''''                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND deltaflag = 'D' "
''''                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND BARCODE = '" & Trim(lsID) & "' "
''''                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
''''                    res = db_select_Col(gLocal, SQL_LOCAL)
''''
''''                    '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
''''                    If CCur(gReadBuf(0)) > 0 Then
''''                        Send_State = "2"
''''                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
''''                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
''''                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
''''                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
''''                        SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
''''                    ElseIf CCur(gReadBuf(0)) = 0 Then
''''                        Send_State = "3"
''''                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
''''                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
''''                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
''''                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
''''                        SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
''''                        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
''''                        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
''''                    End If
''''                End If
'''                '//////////////////
'''
'''                Send_State = "1" '/  <---------- 혈액학장비가 아니라서 상태가 1로만 들어감
'''
'''                '/----------------------------- 결과 상태 넣기
'''                If Send_State = "1" Then
'''
'''                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'''                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
'''                ElseIf Send_State = "2" Then
'''
'''                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'''                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                    SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
'''                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
'''                ElseIf Send_State = "3" Then
'''
'''                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'''                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
'''                End If
'''
'''
'''
'''                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
''''                If gComment_All <> "" Or gComment_Code <> "" Then
''''                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & chrCR & gComment_Code & "' "
''''                End If
'''                '/-----------------------------
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
'''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
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
'''                State_M = Mid(State_GM, State_cnt + 1)
'''
'''
'''                '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
'''                If Trim(State_G) <> "" Then
'''                    SQL = "UPDATE SPSLHRRST "
'''
'''                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
'''                        If Send_State = "1" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "2" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "3" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        End If
'''                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '검사코드"
'''                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
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
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "2" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "3" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        End If
'''                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '검사코드"
'''                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
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
'''                '////////// 접수 테이블
'''                SQL = "UPDATE SPSLMJBDI "
'''                If Send_State = "1" Then
'''                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
'''                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
'''                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''                ElseIf Send_State = "2" Then
'''                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'''                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
'''                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''                ElseIf Send_State = "3" Then
'''                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'''                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
'''                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
'''                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''                End If
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "') "                    '검사코드"
'''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
'''                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
'''                res = SendQuery(gServer, SQL)
'''
'''                If res = -1 Then
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
'''        SQL = "UPDATE SPSLMJBBI "
'''        If Send_State = "1" Then
'''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
'''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
'''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        ElseIf Send_State = "2" Then
'''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
'''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
'''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        ElseIf Send_State = "3" Then
'''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
'''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
'''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        End If
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
'''        res = SendQuery(gServer, SQL)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            cn_Ser.RollbackTrans
'''            Exit Function
'''        End If
'''        '/------------------------------------
'''        'db_Commit gServer
'''        cn_Ser.CommitTrans
'''        Insert_Data = 1
'''    End With
'''End Function

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


Function Insert_Data(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsSpecID        As String
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
    Dim strSpcmCd       As String
    Dim strMICCd        As String
    Dim lsCulSpecNo        As String
    
    Dim strVS_HID           As String
    Dim strVS_PID           As String
    Dim strVS_RECENO        As String
    Dim strVS_SPECIMENID    As String
    Dim strVS_SPECIMENCODE  As String
    Dim strVS_SEQ           As String
    Dim strVS_EXAMCODE      As String
    
    Dim strCOMMENTS         As String
    
On Error GoTo Err

    Insert_Data = -1

    lsID = ""
    lsID = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
    lsSpecNo = Trim(GetText(frmInterface.vasID, argSpcRow, colSpecNo))
    lsPid = Trim(GetText(frmInterface.vasID, argSpcRow, colPID))
    
'    lsID = ""
'    lsID = Trim(GetText(frmInterface.vasResult, argSpcRow, colBarcode))
'    lsSpecID = Trim(GetText(frmInterface.vasResult, argSpcRow, colSpecNo))

    lsID = Format(lsID, "000000000000")
    lsSpecID = Format(lsSpecID, "000000000000")
    
'    lsSpecNo = Trim(GetText(frmInterface.vasResult, argSpcRow, colBarcode))
'    lsCulSpecNo = lsSpecNo + 1
'    lsPid = Trim(GetText(frmInterface.vasResult, argSpcRow, 5))
'    lsInsertTime = Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
    
    'If lsSpecNo = "" Then
    '    Exit Function
    'End If
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    
    SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
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
    i = 0
    cn_Ser.BeginTrans
    
    '서버로 결과값 저장하기
    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
        strExmnCD = Trim(GetText(frmInterface.vasTemp, iRow, 2))
        strSpcmCd = Trim(GetText(frmInterface.vasTemp, iRow, 10))
        strMICCd = Trim(GetText(frmInterface.vasTemp, iRow, 1))
        strMICCd = UCase(strMICCd)
        sResult = Trim(GetText(frmInterface.vasTemp, iRow, 3))
        sCnt = ""
    
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT * "
        SQL = SQL & vbCrLf & "  FROM V_EXAMRES "
        SQL = SQL & vbCrLf & " WHERE VS_SPECIMENID  = '" & lsID & "' "  '검체번호
        SQL = SQL & vbCrLf & "   AND VS_EXAMCODE    = '" & strExmnCD & "' " '검사코드
        SQL = SQL & vbCrLf & "   AND VS_EXSTATE     IN ('1', '2', '4', '5') "    '/VS_EXAMSTATE(0.미접수(검체번호발행), 1.보고대상(결과미완료), 2.보고대상(결과완료), 3.Confirm대상, 4.중간보고, 5.최종보고)
        If ReadSQL_HIS(SQL, ADR_HIS) = False Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        
        If Not ADR_HIS Is Nothing Then
            strVS_HID = Trim(ADR_HIS!VS_HID & "")
            strVS_PID = Trim(ADR_HIS!VS_PID & "")
            strVS_RECENO = Trim(ADR_HIS!VS_RECENO & "")
            strVS_SPECIMENID = Trim(ADR_HIS!VS_SPECIMENID & "")
            strVS_SPECIMENCODE = Trim(ADR_HIS!VS_SPECIMENCODE & "")
            strVS_SEQ = Trim(ADR_HIS!VS_SEQ & "")
            strVS_EXAMCODE = Trim(ADR_HIS!VS_EXAMCODE & "")
            
            ADR_HIS.Close: Set ADR_HIS = Nothing
            
            SQL = ""
            SQL = SQL & vbCrLf & "SELECT * "
            SQL = SQL & vbCrLf & "  FROM AFBSTAINPCR "
            SQL = SQL & vbCrLf & " WHERE SPECIMENID   =   '" & strVS_SPECIMENID & "' "
            SQL = SQL & vbCrLf & "   AND EXAMCODE     =   '" & strVS_EXAMCODE & "' "
            If ReadSQL_HIS(SQL, ADR_HIS) = False Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
            If Not ADR_HIS Is Nothing Then
                ADR_HIS.Close: Set ADR_HIS = Nothing
                
                '/장비값=>변환값(- => 0, + => 1)
                SQL = ""
                SQL = SQL & vbCrLf & "UPDATE AFBSTAINPCR SET "
                Select Case sResult
                    Case "-":
                        strCOMMENTS = "When we performed RT-PCR" & vbCrLf & "the test result suggest no evidence of M. tuberculosis."
                        
                        SQL = SQL & vbCrLf & "       NEGAPOSIFLAG   =   '0', " '/NEGAPOSIFLAG(PCR결과)
                        SQL = SQL & vbCrLf & "       INTERPRETATION =   'Negative (RT-PCR)', "
                        SQL = SQL & vbCrLf & "       COMMENTS       =   '" & strCOMMENTS & "', "

                    Case "+":
                        strCOMMENTS = "When we performed RT-PCR" & vbCrLf & "the test result suggest evidence of M. tuberculosis."
                        
                        SQL = SQL & vbCrLf & "       NEGAPOSIFLAG   =   '1', "
                        SQL = SQL & vbCrLf & "       INTERPRETATION =   'Positive (RT-PCR)', "
                        SQL = SQL & vbCrLf & "       COMMENTS       =   '" & strCOMMENTS & "', "
                        
                    Case Else:
                        SQL = SQL & vbCrLf & "       NEGAPOSIFLAG   =   '', "
                        SQL = SQL & vbCrLf & "       INTERPRETATION =   '', "
                        SQL = SQL & vbCrLf & "       COMMENTS       =   '', "
                End Select
                SQL = SQL & vbCrLf & "       INPUT_UID      =   '" & gUserID & "', "
                SQL = SQL & vbCrLf & "       INPUT_DATETIME =   SYSDATE "
                SQL = SQL & vbCrLf & " WHERE SPECIMENID     =   '" & strVS_SPECIMENID & "' "
                SQL = SQL & vbCrLf & "   AND EXAMCODE       =   '" & strVS_EXAMCODE & "' "
            Else
                SQL = ""
                SQL = SQL & vbCrLf & "INSERT INTO AFBSTAINPCR "
                SQL = SQL & vbCrLf & " (HID,        PID,    RECENO,         SPECIMENID, SPECIMENCODE, "
                SQL = SQL & vbCrLf & "  EXAMCODE,   RESSEQ, NEGAPOSIFLAG,   INTERPRETATION, COMMENTS, INPUT_UID,  INPUT_DATETIME) "
                SQL = SQL & vbCrLf & " VALUES "
                SQL = SQL & vbCrLf & " ('" & strVS_HID & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_PID & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_RECENO & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_SPECIMENID & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_SPECIMENCODE & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_EXAMCODE & "',"
                SQL = SQL & vbCrLf & "   " & Val(strVS_SEQ) & ", "
                Select Case sResult
                    Case "-":
                        strCOMMENTS = "When we performed RT-PCR" & vbCrLf & "the test result suggest no evidence of M. tuberculosis."
                    
                        SQL = SQL & vbCrLf & "  '0', " '/NEGAPOSIFLAG(PCR결과)
                        SQL = SQL & vbCrLf & "  'Negative (RT-PCR)', "
                        SQL = SQL & vbCrLf & "  '" & strCOMMENTS & "', "
                    
                    Case "+":
                        strCOMMENTS = "When we performed RT-PCR" & vbCrLf & "the test result suggest evidence of M. tuberculosis."
                        
                        SQL = SQL & vbCrLf & "  '1', "
                        SQL = SQL & vbCrLf & "  'Positive (RT-PCR)', "
                        SQL = SQL & vbCrLf & "  '" & strCOMMENTS & "', "
                    
                    Case Else:
                        SQL = SQL & vbCrLf & "  '', "
                        SQL = SQL & vbCrLf & "  '', "
                        SQL = SQL & vbCrLf & "  '', "
                End Select
                SQL = SQL & vbCrLf & "  '" & gUserID & "', "
                SQL = SQL & vbCrLf & "  SYSDATE) "
            End If
            res = SendQuery(gServer, SQL)
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        
            SQL = ""
            SQL = SQL & vbCrLf & "UPDATE EXAMRES SET "
            SQL = SQL & vbCrLf & "       EXAMUID    =   '" & gUserID & "', "
            '''SQL = SQL & vbCrLf & "       EXAMDATE   =   SYSDATE, " '/판독일로 쓰이는데 전산과장이 변경되지 않게 해달라고 함.
            SQL = SQL & vbCrLf & "       EXAMSTATE  =   'D' "
            SQL = SQL & vbCrLf & " WHERE SPECIMENID =   '" & strVS_SPECIMENID & "'"
            SQL = SQL & vbCrLf & "   AND EXAMCODE   =   '" & strVS_EXAMCODE & "' "
            res = SendQuery(gServer, SQL)
            
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        End If
        
'/오덕진 막음(2011.12.20)
'''        SQL = "SELECT count(*) FROM EXAMRES "
'''        SQL = SQL & vbCrLf & " WHERE SPECIMENID = '" & lsSpecID & "' "                                                 '검체번호"
'''        SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & strExmnCD & "' "  '검사코드"
'''        SQL = SQL & vbCrLf & "   AND EXAMSTATE = 'B' "                                                          '결과상태"
'''        res = db_select_Col(gServer, SQL)
'''        If res > 0 And gReadBuf(0) = "0" Then
'''                       SQL = "UPDATE EXAMRES "
'''            SQL = SQL & vbCrLf & "   SET RESULT = '" & sResult & "' "
'''            SQL = SQL & vbCrLf & "       ,EXAMUID = '" & gUserID & "' "
'''            SQL = SQL & vbCrLf & "       ,EXAMDATE = SYSDATE "
'''            SQL = SQL & vbCrLf & "       ,EXAMSTATE = "
'''            SQL = SQL & vbCrLf & "                 (CASE "
'''            SQL = SQL & vbCrLf & "                  WHEN NVL(EXAMSTATE,' ')  = 'B' "
'''            SQL = SQL & vbCrLf & "                       THEN 'D'"
'''            SQL = SQL & vbCrLf & "                  END) "
'''            SQL = SQL & vbCrLf & " WHERE RECENO = '" & lsSpecNo & "'"
'''            SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & strExmnCD & "' "
'''            SQL = SQL & vbCrLf & "   AND NVL(RESEND,' ') || NVL(EXAMSTATE,' ') <> '1D' "
'''            SQL = SQL & vbCrLf & "   AND LABRECYN = 'Y' "
'''                                             '결과상태"
'''            res = SendQuery(gServer, SQL)
'''
'''            If res < 0 Then
'''                SaveQuery SQL
'''                cn_Ser.RollbackTrans
'''                Exit Function
'''            End If
'''        End If
'/오덕진 막음(2011.12.20)

    Next iRow
    
    SQL = ""

    cn_Ser.CommitTrans
       
    Insert_Data = 1
    
    Exit Function
    
Err:
    cn_Ser.RollbackTrans
    
    
End Function

''
''Function RsltState_Check(asSpecNo As String, asExamCode As String) As String '/// 결과 형태 : (그룹코드/멀티코드) : 상태가 중간보고 이하일때
''    Dim PRSC_CD_G       As String
''    Dim EXMN_CD         As String
''    Dim PRSC_CD_M       As String
''
''
''    RsltState_Check = ""
''    PRSC_CD_G = " "
''    PRSC_CD_M = " "
''
''    SQL = ""
''    SQL = SQL & vbCrLf & "SELECT DISTINCT "
''    SQL = SQL & vbCrLf & "       R1.PRSC_CD "
''    'SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
''    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
''    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
''    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
''    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
''    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
''    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD = '" & asExamCode & "' "
''    SQL = SQL & vbCrLf & "   AND R1.PRSC_CD LIKE ('%G%') "
'''    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '2' "
''    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
''    res = db_select_Col(gServer, SQL)
''    If gReadBuf(0) <> "" Then: PRSC_CD_G = gReadBuf(0)
''    gReadBuf(0) = ""
''
''    SQL = ""
''    SQL = SQL & vbCrLf & "SELECT DISTINCT "
''    SQL = SQL & vbCrLf & "       R1.PRSC_CD "
''    'SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
''    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
''    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
''    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
''    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
''    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
''    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
''    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('M') "
'''    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '2' "
''    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
''    res = db_select_Col(gServer, SQL)
''
''    If gReadBuf(0) <> "" Then: PRSC_CD_M = gReadBuf(0)
''    gReadBuf(0) = ""
''
''
''    RsltState_Check = PRSC_CD_G & "/" & PRSC_CD_M
''
''End Function


Function RsltState_Check(asSpecNo As String, asExamCode As String) As String '/// 결과 형태 : (그룹코드/멀티코드) : 상태가 중간보고 이하일때
    Dim PRSC_CD_G       As String
    Dim EXMN_CD         As String
    Dim PRSC_CD_M       As String
    Dim PRSC_CD_B       As String
    
    RsltState_Check = ""
    PRSC_CD_G = " "
    PRSC_CD_M = " "
    PRSC_CD_B = " "
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    SQL = SQL & vbCrLf & "       R1.PRSC_CD "
    'SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD = '" & asExamCode & "' "
    SQL = SQL & vbCrLf & "   AND R1.PRSC_CD LIKE ('%G%') "
'    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
    res = db_select_Col(gServer, SQL)
    If gReadBuf(0) <> "" Then: PRSC_CD_G = gReadBuf(0)
    gReadBuf(0) = ""
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    'SQL = SQL & vbCrLf & "       R1.PRSC_CD "
    SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('M') "
'    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
    res = db_select_Col(gServer, SQL)
       
    If gReadBuf(0) <> "" Then: PRSC_CD_M = gReadBuf(0)
    gReadBuf(0) = ""

    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    'SQL = SQL & vbCrLf & "       R1.PRSC_CD "
    SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('B') "
'    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
    res = db_select_Col(gServer, SQL)
       
    If gReadBuf(0) <> "" Then: PRSC_CD_B = gReadBuf(0)
    gReadBuf(0) = ""
    
    
    RsltState_Check = PRSC_CD_G & "/" & PRSC_CD_M & "/" & PRSC_CD_B
    
End Function

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
        
On Error GoTo Err

    Insert_Data_MIC = -1
    
    lsID = ""
    lsID = Trim(GetText(frmInterface.vasResult, argSpcRow, colBarcode))
    lsSpecNo = Trim(GetText(frmInterface.vasResult, argSpcRow, colSpecNo))
    lsPid = Trim(GetText(frmInterface.vasResult, argSpcRow, colPID))
    lsInsertTime = Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
    
    If lsSpecNo = "" Then
        Exit Function
    End If
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    
    SQL = " Select isocd, equipcode, examcode, result, antsize, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
          " And barcode = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, colBarcode)) & "' " & vbCrLf & _
          " And diskno = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, colRack)) & "' " & vbCrLf & _
          " And posno = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, colPos)) & "' "
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
        
        If iRow = 1 Then
            '-- 미생물 세균결과
            SQL = "SELECT SPCM_NO FROM SPSLHMBAC "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "' "                                                    '환자번호"
            SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "' "                                                     '결과상태"
            SQL = SQL & vbCrLf & "   AND BCTR_SQNO = " & iRow                       '검사코드"
            res = db_select_Col(gServer, SQL)
            If res > 0 Then
                SQL = "UPDATE SPSLHMBAC SET "
                SQL = SQL & " SORT_SEQ = '', "
                SQL = SQL & " SPCM_CD = '', "
                SQL = SQL & " CLTR_VOL_CD = '', "
                SQL = SQL & " CLTR_PERD = '', "
                SQL = SQL & " PRE_RSLT_CD = '', "
                SQL = SQL & " MDDL_RPTR_ID = '', "
                SQL = SQL & " LAST_BCTR_CD = '', "
                SQL = SQL & " MDDL_RPTG_DT = '', "
                SQL = SQL & " LAST_RPTR_ID = '', "
                SQL = SQL & " LAST_RPTG_DT = '', "
                SQL = SQL & " RSLT_STAT = '', "
                SQL = SQL & " CMNT_DVSN = '', "
                SQL = SQL & " EQPM_CD = '', "
                SQL = SQL & " RMRK = '', "
                SQL = SQL & " REGI_ID = '', "
                SQL = SQL & " RGST_DT = '', "
                SQL = SQL & " AMEN_ID = '', "
                SQL = SQL & " UPDT_DT = '' "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & lsPid & "' "
                SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 8)) & "' "
                SQL = SQL & vbCrLf & "   AND BCTR_SQNO = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "
                
                res = SendQuery(gServer, SQL)
                
                If res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            Else
                SQL = "INSERT INTO SPSLHMBAC (SPCM_NO, EXMN_CD, BCTR_CD,BCTR_SQNO,"
                SQL = SQL & vbCrLf & "SORT_SEQ,SPCM_CD,CLTR_VOL_CD,CLTR_PERD,PRE_RSLT_CD,MDDL_RPTR_ID,LAST_BCTR_CD,"
                SQL = SQL & vbCrLf & "MDDL_RPTG_DT , LAST_RPTR_ID, LAST_RPTG_DT, RSLT_STAT, CMNT_DVSN, EQPM_CD, RMRK, REGI_ID, RGST_DT, AMEN_ID, UPDT_DT)"
                SQL = SQL & vbCrLf & " Values ( "
                SQL = SQL & vbCrLf & " '" & lsSpecNo & "', "
                SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "
                SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "
                SQL = SQL & vbCrLf & " '" & iRow & "', "
                SQL = SQL & vbCrLf & " '" & iRow & "', "
                SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 5)) & "', "    'spcm_cd
                SQL = SQL & vbCrLf & " '', "    'CLTR_VOL_CD
                SQL = SQL & vbCrLf & " '', "    'CLTR_PERD
                SQL = SQL & vbCrLf & " '', "    'PRE_RSLT_CD
                SQL = SQL & vbCrLf & " '', "    'MDDL_RPTR_ID
                SQL = SQL & vbCrLf & " '', "    'LAST_BCTR_CD
                SQL = SQL & vbCrLf & " '', "    'MDDL_RPTG_DT
                SQL = SQL & vbCrLf & " '', "    'LAST_RPTR_ID
                SQL = SQL & vbCrLf & " '', "    'RSLT_STAT
                SQL = SQL & vbCrLf & " '', "    'CMNT_DVSN
                SQL = SQL & vbCrLf & " '', "    'EQPM_CD
                SQL = SQL & vbCrLf & " '', "    'RMRK
                SQL = SQL & vbCrLf & " '', "    'REGI_ID
                SQL = SQL & vbCrLf & " '', "    'RGST_DT
                SQL = SQL & vbCrLf & " '', "    'AMEN_ID
                SQL = SQL & vbCrLf & " sysdate) "   'UPDT_DT
                res = SendQuery(gServer, SQL)
                
                If res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            
            End If
        End If
        
        '-- 미생물 항생제결과
        SQL = "SELECT SPCM_NO FROM SPSLHMANT "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & lsPid & "' "                                                    '환자번호"
        SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 8)) & "' "                                                     '결과상태"
        SQL = SQL & vbCrLf & "   AND BCTR_SQNO = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '검사코드"
        SQL = SQL & vbCrLf & "   AND ANTB_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '검사코드"
        res = db_select_Col(gServer, SQL)
        If res > 0 Then
            SQL = "UPDATE SPSLHMANT "
                SQL = SQL & " SPCM_CD = '', "
                SQL = SQL & " ANTB_RSLT = '', "
                SQL = SQL & " DTRM_RSLT = '', "
                SQL = SQL & " ANTB_EXMN_MTHD = '', "
                SQL = SQL & " RSLT_RPTR_ID = '', "
                SQL = SQL & " RSLT_RPTG_DT = '', "
                SQL = SQL & " MDDL_RPTG_ID = '', "
                SQL = SQL & " MDDL_RPTG_DT = '', "
                SQL = SQL & " LAST_RPTR_ID = '', "
                SQL = SQL & " LAST_RPTG_DT = '', "
                SQL = SQL & " RSLT_STAT = '', "
                SQL = SQL & " EQPM_CD = '', "
                SQL = SQL & " REGI_ID = '', "
                SQL = SQL & " RGST_DT = '', "
                SQL = SQL & " AMEN_ID = '', "
                SQL = SQL & " UPDT_DT = '' "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & lsPid & "' "
                SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 8)) & "' "
                SQL = SQL & vbCrLf & "   AND BCTR_SQNO = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "
                SQL = SQL & vbCrLf & "   AND ANTB_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "
            res = SendQuery(gServer, SQL)
            
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        Else
            SQL = "INSERT INTO SPSLHMANT (SPCM_NO,EXMN_CD,BCTR_CD,BCTR_SQNO,ANTB_CD,"
            SQL = SQL & vbCrLf & "SPCM_CD,ANTB_RSLT,DTRM_RSLT,ANTB_EXMN_MTHD,RSLT_RPTR_ID,RSLT_RPTG_DT,MDDL_RPTR_ID,MDDL_RPTG_DT,"
            SQL = SQL & vbCrLf & "LAST_RPTR_ID , LAST_RPTG_DT, RSLT_STAT, EQPM_CD, REGI_ID, RGST_DT, AMEN_ID, UPDT_DT)"
            SQL = SQL & vbCrLf & " Values ( "
            SQL = SQL & vbCrLf & " '" & lsSpecNo & "', "
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "
            SQL = SQL & vbCrLf & " '" & iRow & "', "
            SQL = SQL & vbCrLf & " '" & iRow & "', "
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 5)) & "', "    'spcm_cd
            SQL = SQL & vbCrLf & " '', "    'ANTB_RSLT
            SQL = SQL & vbCrLf & " '', "    'DTRM_RSLT
            SQL = SQL & vbCrLf & " '', "    'ANTB_EXMN_MTHD
            SQL = SQL & vbCrLf & " '', "    'RSLT_RPTR_ID
            SQL = SQL & vbCrLf & " '', "    'RSLT_RPTG_DT
            SQL = SQL & vbCrLf & " '', "    'MDDL_RPTR_ID
            SQL = SQL & vbCrLf & " '', "    'MDDL_RPTG_DT
            SQL = SQL & vbCrLf & " '', "    'LAST_RPTR_ID
            SQL = SQL & vbCrLf & " '', "    'LAST_RPTG_DT
            SQL = SQL & vbCrLf & " '', "    'RSLT_STAT
            SQL = SQL & vbCrLf & " '', "    'EQPM_CD
            SQL = SQL & vbCrLf & " '', "    'REGI_ID
            SQL = SQL & vbCrLf & " '', "    'RGST_DT
            SQL = SQL & vbCrLf & " '', "    'AMEN_ID
            SQL = SQL & vbCrLf & " sysdate) "   'UPDT_DT
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
       
    Insert_Data_MIC = 1
    
    Exit Function
    
Err:
    cn_Ser.RollbackTrans
    
    
End Function

Function Insert_Data_R(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsSpecID        As String
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
    Dim strSpcmCd       As String
    Dim strMICCd        As String
    Dim lsCulSpecNo        As String
    
    Dim strVS_HID           As String
    Dim strVS_PID           As String
    Dim strVS_RECENO        As String
    Dim strVS_SPECIMENID    As String
    Dim strVS_SPECIMENCODE  As String
    Dim strVS_SEQ           As String
    Dim strVS_EXAMCODE      As String
    
    Dim strCOMMENTS         As String
    
On Error GoTo Err

    Insert_Data_R = -1

    lsID = ""
    lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, colBarcode))
    lsSpecNo = Trim(GetText(frmInterface.vasRID, argSpcRow, colSpecNo))
    lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, colPID))

    lsID = Format(lsID, "000000000000")
    lsSpecID = Format(lsSpecID, "000000000000")
        
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    
    SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
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
    i = 0
    cn_Ser.BeginTrans
    
    '서버로 결과값 저장하기
    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
        strExmnCD = Trim(GetText(frmInterface.vasTemp, iRow, 2))
        strSpcmCd = Trim(GetText(frmInterface.vasTemp, iRow, 10))
        strMICCd = Trim(GetText(frmInterface.vasTemp, iRow, 1))
        strMICCd = UCase(strMICCd)
        sResult = Trim(GetText(frmInterface.vasTemp, iRow, 3))
        sCnt = ""
    
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT * "
        SQL = SQL & vbCrLf & "  FROM V_EXAMRES "
        SQL = SQL & vbCrLf & " WHERE VS_SPECIMENID  = '" & lsID & "' "  '검체번호
        SQL = SQL & vbCrLf & "   AND VS_EXAMCODE    = '" & strExmnCD & "' " '검사코드
        SQL = SQL & vbCrLf & "   AND VS_EXSTATE     IN ('1', '2', '4', '5') "    '/VS_EXAMSTATE(0.미접수(검체번호발행), 1.보고대상(결과미완료), 2.보고대상(결과완료), 3.Confirm대상, 4.중간보고, 5.최종보고)
        If ReadSQL_HIS(SQL, ADR_HIS) = False Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        
        If Not ADR_HIS Is Nothing Then
            strVS_HID = Trim(ADR_HIS!VS_HID & "")
            strVS_PID = Trim(ADR_HIS!VS_PID & "")
            strVS_RECENO = Trim(ADR_HIS!VS_RECENO & "")
            strVS_SPECIMENID = Trim(ADR_HIS!VS_SPECIMENID & "")
            strVS_SPECIMENCODE = Trim(ADR_HIS!VS_SPECIMENCODE & "")
            strVS_SEQ = Trim(ADR_HIS!VS_SEQ & "")
            strVS_EXAMCODE = Trim(ADR_HIS!VS_EXAMCODE & "")
            
            ADR_HIS.Close: Set ADR_HIS = Nothing
            
            SQL = ""
            SQL = SQL & vbCrLf & "SELECT * "
            SQL = SQL & vbCrLf & "  FROM AFBSTAINPCR "
            SQL = SQL & vbCrLf & " WHERE SPECIMENID   =   '" & strVS_SPECIMENID & "' "
            SQL = SQL & vbCrLf & "   AND EXAMCODE     =   '" & strVS_EXAMCODE & "' "
            If ReadSQL_HIS(SQL, ADR_HIS) = False Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
            If Not ADR_HIS Is Nothing Then
                ADR_HIS.Close: Set ADR_HIS = Nothing
                
                '/장비값=>변환값(- => 0, + => 1)
                SQL = ""
                SQL = SQL & vbCrLf & "UPDATE AFBSTAINPCR SET "
                Select Case sResult
                    Case "-":
                        strCOMMENTS = "When we performed RT-PCR" & vbCrLf & "the test result suggest no evidence of M. tuberculosis."
                        
                        SQL = SQL & vbCrLf & "       NEGAPOSIFLAG   =   '0', " '/NEGAPOSIFLAG(PCR결과)
                        SQL = SQL & vbCrLf & "       INTERPRETATION =   'Negative (RT-PCR)', "
                        SQL = SQL & vbCrLf & "       COMMENTS       =   '" & strCOMMENTS & "', "

                    Case "+":
                        strCOMMENTS = "When we performed RT-PCR" & vbCrLf & "the test result suggest evidence of M. tuberculosis."
                        
                        SQL = SQL & vbCrLf & "       NEGAPOSIFLAG   =   '1', "
                        SQL = SQL & vbCrLf & "       INTERPRETATION =   'Positive (RT-PCR)', "
                        SQL = SQL & vbCrLf & "       COMMENTS       =   '" & strCOMMENTS & "', "
                        
                    Case Else:
                        SQL = SQL & vbCrLf & "       NEGAPOSIFLAG   =   '', "
                        SQL = SQL & vbCrLf & "       INTERPRETATION =   '', "
                        SQL = SQL & vbCrLf & "       COMMENTS       =   '', "
                End Select
                SQL = SQL & vbCrLf & "       INPUT_UID      =   '" & gUserID & "', "
                SQL = SQL & vbCrLf & "       INPUT_DATETIME =   SYSDATE "
                SQL = SQL & vbCrLf & " WHERE SPECIMENID     =   '" & strVS_SPECIMENID & "' "
                SQL = SQL & vbCrLf & "   AND EXAMCODE       =   '" & strVS_EXAMCODE & "' "
            Else
                SQL = ""
                SQL = SQL & vbCrLf & "INSERT INTO AFBSTAINPCR "
                SQL = SQL & vbCrLf & " (HID,        PID,    RECENO,         SPECIMENID, SPECIMENCODE, "
                SQL = SQL & vbCrLf & "  EXAMCODE,   RESSEQ, NEGAPOSIFLAG,   INTERPRETATION, COMMENTS, INPUT_UID,  INPUT_DATETIME) "
                SQL = SQL & vbCrLf & " VALUES "
                SQL = SQL & vbCrLf & " ('" & strVS_HID & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_PID & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_RECENO & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_SPECIMENID & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_SPECIMENCODE & "', "
                SQL = SQL & vbCrLf & "  '" & strVS_EXAMCODE & "',"
                SQL = SQL & vbCrLf & "   " & Val(strVS_SEQ) & ", "
                Select Case sResult
                    Case "-":
                        strCOMMENTS = "When we performed RT-PCR" & vbCrLf & "the test result suggest no evidence of M. tuberculosis."
                    
                        SQL = SQL & vbCrLf & "  '0', " '/NEGAPOSIFLAG(PCR결과)
                        SQL = SQL & vbCrLf & "  'Negative (RT-PCR)', "
                        SQL = SQL & vbCrLf & "  '" & strCOMMENTS & "', "
                    
                    Case "+":
                        strCOMMENTS = "When we performed RT-PCR" & vbCrLf & "the test result suggest evidence of M. tuberculosis."
                        
                        SQL = SQL & vbCrLf & "  '1', "
                        SQL = SQL & vbCrLf & "  'Positive (RT-PCR)', "
                        SQL = SQL & vbCrLf & "  '" & strCOMMENTS & "', "
                    
                    Case Else:
                        SQL = SQL & vbCrLf & "  '', "
                        SQL = SQL & vbCrLf & "  '', "
                        SQL = SQL & vbCrLf & "  '', "
                End Select
                SQL = SQL & vbCrLf & "  '" & gUserID & "', "
                SQL = SQL & vbCrLf & "  SYSDATE) "
            End If
            res = SendQuery(gServer, SQL)
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        
            SQL = ""
            SQL = SQL & vbCrLf & "UPDATE EXAMRES SET "
            SQL = SQL & vbCrLf & "       EXAMUID    =   '" & gUserID & "', "
            '''SQL = SQL & vbCrLf & "       EXAMDATE   =   SYSDATE, " '/판독일로 쓰이는데 전산과장이 변경되지 않게 해달라고 함.
            SQL = SQL & vbCrLf & "       EXAMSTATE  =   'D' "
            SQL = SQL & vbCrLf & " WHERE SPECIMENID =   '" & strVS_SPECIMENID & "'"
            SQL = SQL & vbCrLf & "   AND EXAMCODE   =   '" & strVS_EXAMCODE & "' "
            res = SendQuery(gServer, SQL)
            
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        End If
        

    Next iRow
    
    SQL = ""

    cn_Ser.CommitTrans
       
    Insert_Data_R = 1
    
    Exit Function
    
Err:
    cn_Ser.RollbackTrans

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

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    
    Get_Sample_Info = -1
    '환자정보 가져오기
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '환자번호, 환자이름, 주민번호, 성별, 나이
    '-- 검사대상자 가져오기
          SQL = "Select Distinct a.PID, b.PNAME, b.SEX, a.RECENO, a.SEQNO, a.EXAMCODE, a.SPECIMENCODE,a.SPECIMENID "
    SQL = SQL & "  From EXAMRES a, PATIENT b"
'    SQL = SQL & " Where a.EXAMSTATE = 'B' "   '접수
'    SQL = SQL & "   AND a.PID = b.PID "   '접수
'    SQL = SQL & "   AND (NVL(a.RESEND,' ') <> '1' "
'    SQL = SQL & "        OR (a.RESEND = '1' AND a.EXAMSTATE = 'E')) "
'    SQL = SQL & "   AND a.SPECIMENID = '" & Format(sBarcode, "000000000000") & "'"
    
    SQL = SQL & " Where a.EXAMSTATE >= 'B' "   '접수
    SQL = SQL & "   AND a.PID = b.PID "   '접수
    
    SQL = SQL & "   AND a.SPECIMENID = '" & Format(sBarcode, "000000000000") & "'"
    'SQL = SQL & "   AND a.RECENO = '" & sBarcode & "'"
    
    res = db_select_Col(gServer, SQL)
    
    '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
    
    If res = 1 Then
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colSpecNo     '2
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID    '6
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName  '7
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '8
        SetText frmInterface.vasID, "", asRow, colAge    '9
        SetText frmInterface.vasID, Trim(gReadBuf(5)), asRow, colRack    '4
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
    
    '-- 검사코드 가져오기
    SQL = " Select EXAMCODE From EXAMRES " & CR & _
          " Where RECENO = '" & Trim(argEquipCode) & "' " & vbCrLf & _
          "   and RECENO IS NOT NULL"
    
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


