Attribute VB_Name = "DB_Query"
Option Explicit

Public Const colCheckBox = 1
Public Const colSpecNo = 2
Public Const colBarcode = 3
Public Const colRack = 4
Public Const colPos = 5
Public Const colPID = 6
Public Const colPName = 7
Public Const colSex = 8
Public Const colAge = 9
Public Const colOCnt = 10
Public Const colRCnt = 11
Public Const colState = 12
Public Const colHct = 13
Public Const colA1c = 13
Public Const colIFCC = 15
Public Const coleAg = 17


'sendflag
'0: Order
'1: Result
'2: Trans

'vasres, vasrres colum
Public Const colEquipCode = 1
Public Const colExamCode = 2
Public Const colExamName = 3
Public Const colResult = 4
Public Const colSeq = 5
Public Const colFLAG = 6
Public Const colEquipResult = 7
Public Const colDelta = 8
Public Const colPanic = 9

'장비코드로 검사코드 찾기
Public gEquipExamCode As String

'해당검사에 대한 소수점
Public gExamRange As String
'참고치 및 검사명
Public gExamName As String
Public gRFVL_DVSN As String
Public gMALE_HIGH As String
Public gMALE_LOW As String
Public gFEML_HIGH As String
Public gFEML_LOW As String
Public gDELT_DVSN As String
Public gDELT_HIGH As String
Public gDELT_LOW As String
Public gDELT_DD As String
Public gPANC_DVSN As String
Public gPANC_HIGH As String
Public gPANC_LOW As String

Public gTLA_Equip As String
Public gTLA_Sub1 As String
Public gTLA_Sub2 As String

'////Lasc
Public gEXAM_CBC        As String
Public gEXAM_Diff       As String
Public gEXAM_Reti       As String
Public gEXAM_CBC_Diff   As String

'////Comment
Public gComment_All As String
Public gComment_Code As String


Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As Long = 0)
    Dim sCnt As String
    Dim sExamDate As String
    Dim RCnt As Integer
    Dim OCnt As Integer
    
'    SQL = "SELECT COUNT(*) FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
'          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
'    res = db_select_Col(gLocal, SQL)
    With frmInterface
        sExamDate = Format(.dtpToday, "yyyymmdd")
        
        SQL = "DELETE FROM PAT_RES " & vbCrLf & _
              "WHERE EXAMDATE = '" & Format(.dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
              "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              "  AND BARCODE = '" & Trim(GetText(.vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(GetText(.vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
              "  and examcode= '" & Trim(GetText(.vasRes, asRow2, colExamCode)) & "'"
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        
        SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
              "POSNO, PID, PNAME, " & vbCrLf & _
              "PSEX, PAGE, " & vbCrLf & _
              "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
              "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, EQUIPRESULT, RECENO, RESFLAG) " & vbCrLf & _
              "VALUES('" & gEquip & "', '" & Trim(GetText(.vasID, asRow1, colBarcode)) & "', '" & Trim(GetText(.vasID, asRow1, colRack)) & "', " & vbCrLf & _
              "'" & Trim(GetText(.vasID, asRow1, colPos)) & "', '" & Trim(GetText(.vasID, asRow1, colPID)) & "', '" & Trim(GetText(.vasID, asRow1, colPName)) & "', " & vbCrLf & _
              "'" & Trim(GetText(.vasID, asRow1, colSex)) & "', " & 0 & ", " & vbCrLf & _
              "'" & Trim(sExamDate) & "', '" & Trim(GetText(.vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(.vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
              "'" & Trim(GetText(.vasRes, asRow2, colSeq)) & "', '" & Trim(GetText(.vasRes, asRow2, colResult)) & "', '" & Trim(GetText(.vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
              "'" & asSend & "', '" & Trim(GetText(.vasRes, asRow2, 7)) & "', '" & Trim(asEquipResult) & "', '" & Trim(GetText(.vasID, asRow1, colSpecNo)) & "', '') "
        res = SendQuery(gLocal, SQL)

    End With
End Function
''''''////////////////////////// 이전 결과 저장 (2011.10.11) - 효준
''''''Function Insert_Data(ByVal argSpcRow As Integer) As Integer
''''''    Dim iRow            As Integer
''''''    Dim i               As Integer
''''''    Dim j               As Integer
''''''    Dim lsID            As String
''''''    Dim lsSpecNo        As String
''''''    Dim lsPid           As String
''''''    Dim sResult         As String
''''''    Dim sCnt            As String
''''''    Dim sResult1        As String
''''''    Dim sResult2        As String
''''''    Dim ExamCnt         As String
''''''    Dim ExamCode_Spec   As String
''''''    Dim ExamCode_Remark     As String
''''''
''''''    With frmInterface
''''''        gComment_All = ""
''''''        Insert_Data = -1
''''''        ExamCode_Spec = ""
''''''        ExamCode_Remark = ""
''''''        lsID = ""
''''''        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
''''''        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
''''''        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
''''''
''''''        'Local에서 환자별로 결과값 가져오기
''''''        ClearSpread .vasTemp
''''''
''''''        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
''''''              " From pat_res " & vbCrLf & _
''''''              " Where equipno = '" & gEquip & "' " & vbCrLf & _
''''''              " And examdate = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
''''''              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
''''''              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
''''''              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
''''''        res = db_select_Vas(gLocal, SQL, .vasTemp)
''''''
''''''        If res = -1 Then
''''''            SaveQuery SQL
''''''            Exit Function
''''''        End If
''''''
''''''        For i = 1 To frmInterface.vasTemp.DataRowCnt
''''''            If ExamCode_Spec <> "" Then
''''''                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
''''''            Else
''''''                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
''''''            End If
''''''        Next i
''''''
''''''        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
''''''        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
''''''
''''''        gHIVPosFlag = -1
''''''
''''''        sCnt = ""
''''''        sResult1 = ""
''''''        sResult2 = ""
''''''
''''''        SQL = "SELECT EXMN_CD "
''''''        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
''''''        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
''''''        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
''''''        res = db_select_Col(gServer, SQL)
''''''
''''''        j = 0
''''''        Do While gReadBuf(j) <> ""
''''''            If ExamCode_Remark <> "" Then
''''''                ExamCode_Remark = ExamCode_Remark & ",'" & gReadBuf(j) & "'"
''''''            Else
''''''                ExamCode_Remark = "'" & gReadBuf(j) & "'"
''''''            End If
''''''            j = j + 1
''''''        Loop
''''''        If ExamCode_Remark = "" Then ExamCode_Remark = "''"
''''''
''''''        For i = 1 To frmInterface.vasTemp.DataRowCnt
''''''            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
''''''        Next i
''''''
''''''
''''''        cn_Ser.BeginTrans
''''''        '서버로 결과값 저장하기
''''''        For iRow = 1 To .vasTemp.DataRowCnt
''''''
''''''            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
''''''            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
''''''
''''''            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
''''''                gComment_Code = ""
''''''
''''''
''''''                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
''''''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
''''''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
''''''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
''''''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
''''''                res = db_select_Col(gServer, SQL)
''''''
''''''                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
''''''
''''''                sCnt = CLng(gReadBuf(0)) + 1
''''''
''''''
''''''                Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
''''''
''''''
''''''                               SQL = "UPDATE SPSLHRRST "
''''''                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
''''''                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
''''''                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
''''''                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
''''''                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
''''''                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
''''''                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
''''''                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '중간보고자"
''''''                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '중간보고일시"
''''''                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '최종보고자"
''''''                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '최종보고일시"
''''''                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
''''''                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
''''''                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
''''''                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
''''''                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
''''''                If gComment_All <> "" Or gComment_Code <> "" Then
''''''                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & chrCR & gComment_Code & "' "
''''''                End If
''''''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
''''''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
''''''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
''''''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
''''''                res = SendQuery(gServer, SQL)
''''''                If res < 0 Then
''''''                    SaveQuery SQL
''''''                   ' db_RollBack gServer
''''''                   cn_Ser.RollbackTrans
''''''                    Exit Function
''''''                End If
''''''
''''''
''''''
''''''
''''''
''''''                SQL = "UPDATE SPSLMJBDI "
''''''                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
''''''                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
''''''                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
''''''                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
''''''                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''''''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
''''''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
''''''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
''''''                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
''''''                res = SendQuery(gServer, SQL)
''''''
''''''                If res = -1 Then
''''''                    SaveQuery SQL
''''''                    cn_Ser.RollbackTrans
''''''                    Exit Function
''''''                End If
''''''
''''''                SQL = "UPDATE SPSLHRRST "
''''''                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1' "
''''''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
''''''                SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "                     '검사코드"
''''''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
''''''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
''''''                res = SendQuery(gServer, SQL)
''''''
''''''
''''''                If res = -1 Then
''''''                    SaveQuery SQL
''''''                    cn_Ser.RollbackTrans
''''''                    Exit Function
''''''                End If
''''''            End If
''''''        Next iRow
''''''
''''''
''''''
''''''        '//// 결과테이블에서 그룹코드를 제외한 결과중 빈값이 있는경우 처방/접수 테이블에 업데이트 안함
''''''        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
''''''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
''''''        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
''''''        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
''''''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
''''''        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
''''''        res = db_select_Vas(gServer, SQL, .vasTemp1)
''''''        If gReadBuf(0) = "" Then gReadBuf(0) = "0"
''''''        ExamCnt = gReadBuf(0)
''''''        gReadBuf(0) = "0"
''''''
''''''        If ExamCnt = "0" Then                                                         '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외) 업데이트
''''''
''''''            '///////// 처방테이블
''''''            SQL = "UPDATE SPSLMJBBI "
''''''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
''''''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
''''''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''''''            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
''''''            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
''''''            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
''''''            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
''''''            res = SendQuery(gServer, SQL)
''''''
''''''            If res = -1 Then
''''''                SaveQuery SQL
''''''                cn_Ser.RollbackTrans
''''''                Exit Function
''''''            End If
''''''            '////////// 접수 테이블
''''''            SQL = "UPDATE SPSLMJBDI "
''''''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
''''''            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
''''''            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
''''''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
''''''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''''''            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
''''''            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '검사코드"
''''''            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
''''''            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
''''''            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
''''''            res = SendQuery(gServer, SQL)
''''''
''''''            If res = -1 Then
''''''                SaveQuery SQL
''''''                cn_Ser.RollbackTrans
''''''                Exit Function
''''''            End If
''''''
''''''
''''''        ElseIf res = -1 Then                                                             '///// 쿼리 에러인경우
''''''            SaveQuery SQL
''''''            cn_Ser.RollbackTrans
''''''            Exit Function
''''''        Else                                                                             '///// 결과가 미입력일때는 업데이트 안함
''''''
''''''        End If
''''''
''''''        SQL = ""
''''''
''''''
''''''        'db_Commit gServer
''''''        cn_Ser.CommitTrans
''''''        Insert_Data = 1
''''''    End With
''''''End Function

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
    Dim ExamCode_Remark     As String
    
    Dim State_GM    As String       '//// 그룹/멀티 코드
    Dim State_cnt   As Integer      '//// 그룹/멀티 코드 쪽 변수
    Dim State_G     As String       '//// 그룹코드
    Dim State_M     As String       '//// 멀티코드
    Dim State_B     As String       '//// 배터리코드
    
    Dim Send_State      As String
    Dim SQL_LOCAL As String
    

    With frmInterface
        gComment_All = ""
        Insert_Data = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        
        State_GM = ""
        State_cnt = 0
        State_G = ""
        State_M = ""
        lsID = ""
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))

        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        ClearSpread .vasTemp1
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
        
        
        
        '/-------------------------------리마크 처리 때문에 인터페이스에 저장된 코드로 검체를 조회해서 리마크 표시해줄것을 찾음(필요한장비만 열기)
        SQL = "SELECT EXMN_CD "
        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
    
    
        For i = 1 To frmInterface.vasTemp1.DataRowCnt    '/// 실제 검사한 검사코드들
            If ExamCode_Remark <> "" Then
                ExamCode_Remark = ExamCode_Remark & ",'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
            Else
                ExamCode_Remark = "'" & Trim(GetText(frmInterface.vasTemp, i, 1)) & "'"
            End If
        Next i
        
        If ExamCode_Remark = "" Then ExamCode_Remark = "''"

        For i = 1 To frmInterface.vasTemp.DataRowCnt
            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
        Next i
        '/--------------------------------------------------------------------------------------------------------------
        
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If InStr(sResult1, "<") > 0 Then
                sResult1 = Trim(Mid(sResult1, InStr(sResult1, "<") + 1))
            ElseIf InStr(sResult1, ">") > 0 Then
                sResult1 = Trim(Mid(sResult1, InStr(sResult1, ">") + 1))
            End If
            
            If InStr(sResult2, "<") > 0 Then
                sResult2 = "< " & Trim(Mid(sResult2, InStr(sResult2, "<") + 1))
            ElseIf InStr(sResult2, ">") > 0 Then
                sResult2 = "> " & Trim(Mid(sResult2, InStr(sResult2, ">") + 1))
            End If
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                gComment_Code = ""
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
                
                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
                Call Make_Remark(Trim(GetText(.vasTemp, iRow, 2)), Trim(GetText(frmInterface.vasTemp, i, 8)), sResult2)
                '/-----------------------------
                
                               SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                
                
                '/////////// 혈액학 장비만 사용 ( 다른장비들은 결과 입력상태(= 1)로 함)
'                If Mid(Trim(GetText(.vasTemp, iRow, 2)), 1, 2) = "L8" Then
'                    Send_State = "1"
'                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
'                Else
'                    SQL_LOCAL = ""
'                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "SELECT COUNT(EXAMCODE) FROM PAT_RES "
'                    SQL_LOCAL = SQL_LOCAL & vbCrLf & " WHERE (REFFLAG <> '' OR PANICFLAG <> '' OR  DELTAFLAG <> '' ) "
'                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND panicflag = 'P' "
'                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND deltaflag = 'D' "
'                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND BARCODE = '" & Trim(lsID) & "' "
'                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
'                    res = db_select_Col(gLocal, SQL_LOCAL)
'
'                    '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
'                    If CCur(gReadBuf(0)) > 0 Then
'                        Send_State = "2"
'                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'                        SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
'                    ElseIf CCur(gReadBuf(0)) = 0 Then
'                        Send_State = "3"
'                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'                        SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
'                        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'                        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
'                    End If
'                End If
                '//////////////////
                
                Send_State = "1" '/  <---------- 혈액학장비가 아니라서 상태가 1로만 들어감
                
                '/----------------------------- 결과 상태 넣기
                If Send_State = "1" Then

                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
                ElseIf Send_State = "2" Then

                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
                ElseIf Send_State = "3" Then

                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                     '중간보고자"
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                    SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
                End If
                
                '/----------------------------- 결과 상태 넣기
                
                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
                If gComment_All <> "" Or gComment_Code <> "" Then
                    If gComment_All = "" Then
                        SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_Code & "' "
                    Else
                        SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & vbCrLf & gComment_Code & "' "
                    End If
                End If
                '/-----------------------------
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 2)))
                
                State_cnt = InStr(1, State_GM, "/")
                State_G = Mid(State_GM, 1, State_cnt - 1)
                State_GM = Mid(State_GM, State_cnt + 1)
                State_cnt = InStr(1, State_GM, "/")
                State_M = Mid(State_GM, 1, State_cnt - 1)
                State_B = Mid(State_GM, State_cnt + 1)

                
                    
                '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
                If Trim(State_G) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
                '/------------------------------------
                
                '/------------------------------------ 결과테이블 멀티코드 상태 업데이트
                If Trim(State_M) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_M) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
            '/------------------------------------
            
            '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
                If Trim(State_B) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_B) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
            '/------------------------------------
            
            '/------------------------------------ 접수테이블 STATE 업데이트
                '////////// 접수 테이블
                SQL = "UPDATE SPSLMJBDI "
                If Send_State = "1" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf Send_State = "2" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf Send_State = "3" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "','" & Trim(State_B) & "', '" & Trim(GetText(.vasTemp, iRow, 2)) & "') "                    '검사코드"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
        
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If

            '/------------------------------------
            End If
        Next iRow
        
        If Send_State = "" Then cn_Ser.RollbackTrans:   Exit Function
        
        '/------------------------------------ 처방테이블 STATE 업데이트
        '///////// 처방테이블
        SQL = "UPDATE SPSLMJBBI "
        If Send_State = "1" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        ElseIf Send_State = "2" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        ElseIf Send_State = "3" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        End If
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
        res = SendQuery(gServer, SQL)

        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        '/------------------------------------
        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data = 1
    End With
End Function

'//////////////결과 저장 바꿈 (2011.10.11) - 효준
Function Insert_Data_R(ByVal argSpcRow As Integer) As Integer
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
    Dim ExamCode_Remark     As String
    
    Dim State_GM    As String       '//// 그룹/멀티 코드
    Dim State_cnt   As Integer      '//// 그룹/멀티 코드 쪽 변수
    Dim State_G     As String       '//// 그룹코드
    Dim State_M     As String       '//// 멀티코드
    Dim State_B     As String       '//// 배터리코드
    
    Dim Send_State      As String
    Dim SQL_LOCAL As String
    

    With frmInterface
        gComment_All = ""
        Insert_Data_R = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        
        State_GM = ""
        State_cnt = 0
        State_G = ""
        State_M = ""
        lsID = ""
        lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))

        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
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
        
        '/-------------------------------리마크 처리 때문에 인터페이스에 저장된 코드로 검체를 조회해서 리마크 표시해줄것을 찾음(필요한장비만 열기)
        SQL = "SELECT EXMN_CD "
        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
    
    
        For i = 1 To frmInterface.vasTemp1.DataRowCnt    '/// 실제 검사한 검사코드들
            If ExamCode_Remark <> "" Then
                ExamCode_Remark = ExamCode_Remark & ",'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
            Else
                ExamCode_Remark = "'" & Trim(GetText(frmInterface.vasTemp, i, 1)) & "'"
            End If
        Next i
        
        If ExamCode_Remark = "" Then ExamCode_Remark = "''"

        For i = 1 To frmInterface.vasTemp.DataRowCnt
            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
        Next i
        '/--------------------------------------------------------------------------------------------------------------
        
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                gComment_Code = ""
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
                
                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
                Call Make_Remark(Trim(GetText(.vasTemp, iRow, 2)), Trim(GetText(frmInterface.vasTemp, i, 8)), sResult2)
                '/-----------------------------
                
                               SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                
                
                '/////////// 혈액학 장비만 사용 ( 다른장비들은 결과 입력상태(= 1)로 함)
'                If Mid(Trim(GetText(.vasTemp, iRow, 2)), 1, 2) = "L8" Then
'                    Send_State = "1"
'                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
'                Else
'                    SQL_LOCAL = ""
'                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "SELECT COUNT(EXAMCODE) FROM PAT_RES "
'                    SQL_LOCAL = SQL_LOCAL & vbCrLf & " WHERE (REFFLAG <> '' OR PANICFLAG <> '' OR  DELTAFLAG <> '' ) "
'                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND panicflag = 'P' "
'                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND deltaflag = 'D' "
'                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND BARCODE = '" & Trim(lsID) & "' "
'                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
'                    res = db_select_Col(gLocal, SQL_LOCAL)
'
'                    '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
'                    If CCur(gReadBuf(0)) > 0 Then
'                        Send_State = "2"
'                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'                        SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
'                    ElseIf CCur(gReadBuf(0)) = 0 Then
'                        Send_State = "3"
'                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'                        SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
'                        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'                        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
'                    End If
'                End If
                '//////////////////
                
                Send_State = "1" '/  <---------- 혈액학장비가 아니라서 상태가 1로만 들어감
                
                '/----------------------------- 결과 상태 넣기
                If Send_State = "1" Then

                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
                ElseIf Send_State = "2" Then

                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                    SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
                ElseIf Send_State = "3" Then

                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
                End If
                
                
                
                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
                If gComment_All <> "" Or gComment_Code <> "" Then
                    If gComment_All = "" Then
                        SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_Code & "' "
                    Else
                        SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & vbCrLf & gComment_Code & "' "
                    End If
                End If
                '/-----------------------------
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 2)))
                
                State_cnt = InStr(1, State_GM, "/")
                State_G = Mid(State_GM, 1, State_cnt - 1)
                State_GM = Mid(State_GM, State_cnt + 1)
                State_cnt = InStr(1, State_GM, "/")
                State_M = Mid(State_GM, 1, State_cnt - 1)
                State_B = Mid(State_GM, State_cnt + 1)
                    
                '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
                If Trim(State_G) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
                '/------------------------------------
                
                '/------------------------------------ 결과테이블 멀티코드 상태 업데이트
                If Trim(State_M) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
            '/------------------------------------
            
            '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
                If Trim(State_B) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_B) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
            '/------------------------------------
            
            '/------------------------------------ 접수테이블 STATE 업데이트
                If Send_State = "" Then cn_Ser.RollbackTrans: Exit Function
                '////////// 접수 테이블
                SQL = "UPDATE SPSLMJBDI "
                If Send_State = "1" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf Send_State = "2" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf Send_State = "3" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "','" & Trim(State_B) & "', '" & Trim(GetText(.vasTemp, iRow, 2)) & "') "
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
        
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If

            '/------------------------------------
            End If
        Next iRow
        
        '/------------------------------------ 처방테이블 STATE 업데이트
        If Send_State = "" Then cn_Ser.RollbackTrans: Exit Function
        '///////// 처방테이블
        SQL = "UPDATE SPSLMJBBI "
        If Send_State = "1" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        ElseIf Send_State = "2" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        ElseIf Send_State = "3" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        End If
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
        res = SendQuery(gServer, SQL)

        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        '/------------------------------------
        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_R = 1
    End With
End Function

Function Insert_Data_URI(ByVal argSpcRow As Integer) As Integer
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
    Dim ExamCode_Remark     As String
    
    Dim State_GM    As String       '//// 그룹/멀티 코드
    Dim State_cnt   As Integer      '//// 그룹/멀티 코드 쪽 변수
    Dim State_G     As String       '//// 그룹코드
    Dim State_M     As String       '//// 멀티코드
    
    Dim Send_State      As String
    Dim SQL_LOCAL As String
    
    Dim Urin_AutoCode As String     '//// 뇨검사중 자동으로 확인전송으로 넘겨야 할코드들 ( Color, Turbidity)
    
    
    With frmInterface
        gComment_All = ""
        Insert_Data_URI = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        Urin_AutoCode = ""
        
        State_GM = ""
        State_cnt = 0
        State_G = ""
        State_M = ""
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
        
        
        SQL = "SELECT EXAMCODE "
        SQL = SQL & vbCrLf & "FROM EQUIPEXAM "
        SQL = SQL & vbCrLf & "WHERE EQUIPCODE IN ('Col','Tur') "
        res = db_select_Vas(gLocal, SQL, .vasTemp1)
        

        For i = 1 To frmInterface.vasTemp1.DataRowCnt
            If Urin_AutoCode <> "" Then
                Urin_AutoCode = Urin_AutoCode & ",'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
            Else
                Urin_AutoCode = "'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
            End If
        Next i
        
        
        
        '/-------------------------------리마크 처리 때문에 인터페이스에 저장된 코드로 검체를 조회해서 리마크 표시해줄것을 찾음(필요한장비만 열기)
'        SQL = "SELECT EXMN_CD "
'        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
'        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
'        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
'        res = db_select_Col(gServer, SQL)
'
'        j = 0
'        Do While gReadBuf(j) <> ""
'            If ExamCode_Remark <> "" Then
'                ExamCode_Remark = ExamCode_Remark & ",'" & gReadBuf(j) & "'"
'            Else
'                ExamCode_Remark = "'" & gReadBuf(j) & "'"
'            End If
'            j = j + 1
'        Loop
'        If ExamCode_Remark = "" Then ExamCode_Remark = "''"
'
'        For i = 1 To frmInterface.vasTemp.DataRowCnt
'            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
'        Next i
        '/--------------------------------------------------------------------------------------------------------------
        
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                gComment_Code = ""
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1

                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
                'Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
                '/-----------------------------
                
                               SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                
                
                '/////////// 혈액학 장비만 사용 ( 다른장비들은 결과 입력상태(= 1)로 함)
'                If Mid(Trim(GetText(.vasTemp, iRow, 2)), 1, 2) = "L8" Then
'                    Send_State = "1"
'                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
'                Else
'                    SQL_LOCAL = ""
'                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "SELECT COUNT(EXAMCODE) FROM PAT_RES "
'                    SQL_LOCAL = SQL_LOCAL & vbCrLf & " WHERE (REFFLAG <> '' OR PANICFLAG <> '' OR  DELTAFLAG <> '' ) "
'                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND panicflag = 'P' "
'                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND deltaflag = 'D' "
'                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND BARCODE = '" & Trim(lsID) & "' "
'                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
'                    res = db_select_Col(gLocal, SQL_LOCAL)
'
'                    '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
'                    If CCur(gReadBuf(0)) > 0 Then
'                        Send_State = "2"
'                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'                        SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
'                    ElseIf CCur(gReadBuf(0)) = 0 Then
'                        Send_State = "3"
'                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
'                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
'                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'                        SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
'                        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'                        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
'                    End If
'                End If
                '//////////////////
                
                Send_State = "1" '/  <---------- 혈액학장비가 아니라서 상태가 1로만 들어감
                
                '/----------------------------- 결과 상태 넣기
                If Send_State = "1" Then

                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
                ElseIf Send_State = "2" Then

                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                    SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
                ElseIf Send_State = "3" Then

                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
                End If
                
                
                
                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
'                If gComment_All <> "" Or gComment_Code <> "" Then
'                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & chrCR & gComment_Code & "' "
'                End If
                '/-----------------------------
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 2)))
                
                State_cnt = InStr(1, State_GM, "/")
                State_G = Mid(State_GM, 1, State_cnt - 1)
                State_M = Mid(State_GM, State_cnt + 1)
                    
                    
                '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
                If Trim(State_G) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
                '/------------------------------------
                
                '/------------------------------------ 결과테이블 멀티코드 상태 업데이트
                If Trim(State_M) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_M) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
            '/------------------------------------
            
            '/------------------------------------ 접수테이블 STATE 업데이트
                '////////// 접수 테이블
                SQL = "UPDATE SPSLMJBDI "
                If Send_State = "1" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf Send_State = "2" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf Send_State = "3" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "') "                    '검사코드"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
        
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If

            '/------------------------------------
            End If
        Next iRow
        
        '/------------------------------------ 결과테이블 STATE 업데이트(color, turbidity)
        SQL = "UPDATE SPSLHRRST "
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
        SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
        SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & Urin_AutoCode & ") "                                        '검사코드"
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
        res = SendQuery(gServer, SQL)

        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        '/------------------------------------
        
        '/------------------------------------ 처방테이블 STATE 업데이트
        '///////// 처방테이블
        SQL = "UPDATE SPSLMJBBI "
        If Send_State = "1" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        ElseIf Send_State = "2" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        ElseIf Send_State = "3" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        End If
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
        res = SendQuery(gServer, SQL)

        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        '/------------------------------------
        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_URI = 1
    End With
End Function

Function Insert_Data_VAR(ByVal argSpcRow As Integer) As Integer
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
    Dim ExamCode_Remark     As String

    With frmInterface
        gComment_All = ""
        gComment_Code = ""
        Insert_Data_VAR = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
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
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt
            If ExamCode_Spec <> "" Then
                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            Else
                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            End If
        Next i
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt
            If Trim(GetText(frmInterface.vasTemp, i, 1)) <> "A1c" Then
                If gComment_Code <> "" Then
                    gComment_Code = gComment_Code & vbCrLf & Trim(GetText(frmInterface.vasTemp, i, 1)) & " : " & Trim(GetText(frmInterface.vasTemp, i, 3))
                Else
                    gComment_Code = Trim(GetText(frmInterface.vasTemp, i, 1)) & " : " & Trim(GetText(frmInterface.vasTemp, i, 3))
                End If
            End If
        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        

        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" And Trim(GetText(.vasTemp, iRow, 1)) = "A1c" Then
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
                        
                        SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '중간보고자"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '중간보고일시"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '최종보고자"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '최종보고일시"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
                If gComment_All <> "" Or gComment_Code <> "" Then
                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_Code & "' "
                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLMJBDI "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
                
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1' "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                res = SendQuery(gServer, SQL)
                
                
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            Else
            
            End If
        Next iRow
        
        
        
        '//// 결과테이블에서 그룹코드를 제외한 결과중 빈값이 있는경우 처방/접수 테이블에 업데이트 안함
        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        If gReadBuf(0) = "" Then gReadBuf(0) = "0"
        ExamCnt = gReadBuf(0)
        gReadBuf(0) = "0"
        
        If ExamCnt = "0" Then                                                         '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외) 업데이트
            
            '///////// 처방테이블
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            '////////// 접수 테이블
            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '검사코드"
            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
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
        Else                                                                             '///// 결과가 미입력일때는 업데이트 안함
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_VAR = 1
    End With
End Function

Function Insert_Data_R_VAR(ByVal argSpcRow As Integer) As Integer
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
    Dim ExamCode_Remark     As String

    With frmInterface
        gComment_All = ""
        gComment_Code = ""
        Insert_Data_R_VAR = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
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
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt
            If ExamCode_Spec <> "" Then
                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            Else
                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            End If
        Next i
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt
            If Trim(GetText(frmInterface.vasTemp, i, 1)) <> "A1c" Then
                If gComment_Code <> "" Then
                    gComment_Code = gComment_Code & vbCrLf & Trim(GetText(frmInterface.vasTemp, i, 1)) & " : " & Trim(GetText(frmInterface.vasTemp, i, 3))
                Else
                    gComment_Code = Trim(GetText(frmInterface.vasTemp, i, 1)) & " : " & Trim(GetText(frmInterface.vasTemp, i, 3))
                End If
            End If
        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        

        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" And Trim(GetText(.vasTemp, iRow, 1)) = "A1c" Then
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
                        
                        SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '중간보고자"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '중간보고일시"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '최종보고자"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '최종보고일시"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
                If gComment_All <> "" Or gComment_Code <> "" Then
                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_Code & "' "
                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLMJBDI "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
                
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1' "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                res = SendQuery(gServer, SQL)
                
                
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            Else
            
            End If
        Next iRow
        
        
        
        '//// 결과테이블에서 그룹코드를 제외한 결과중 빈값이 있는경우 처방/접수 테이블에 업데이트 안함
        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        If gReadBuf(0) = "" Then gReadBuf(0) = "0"
        ExamCnt = gReadBuf(0)
        gReadBuf(0) = "0"
        
        If ExamCnt = "0" Then                                                         '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외) 업데이트
            
            '///////// 처방테이블
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            '////////// 접수 테이블
            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '검사코드"
            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
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
        Else                                                                             '///// 결과가 미입력일때는 업데이트 안함
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_R_VAR = 1
    End With
End Function


Function Insert_Data_POCT(ByVal argSpcRow As Integer) As Integer
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
    Dim ExamCode_Remark     As String

    With frmInterface
        gComment_All = ""
        Insert_Data_POCT = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
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
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt
            If ExamCode_Spec <> "" Then
                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            Else
                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            End If
        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        
'        SQL = "SELECT EXMN_CD "
'        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
'        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
'        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
'        res = db_select_Col(gServer, SQL)
'
'        j = 0
'        Do While gReadBuf(j) <> ""
'            If ExamCode_Remark <> "" Then
'                ExamCode_Remark = ExamCode_Remark & ",'" & gReadBuf(j) & "'"
'            Else
'                ExamCode_Remark = "'" & gReadBuf(j) & "'"
'            End If
'            j = j + 1
'        Loop
'        If ExamCode_Remark = "" Then ExamCode_Remark = "''"
'
'        For i = 1 To frmInterface.vasTemp.DataRowCnt
'            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
'        Next i
        

        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                gComment_Code = ""
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
                
                
                'Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
                
                
                               SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                    '중간보고자"
                SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '중간보고일시"
                SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                    '최종보고자"
                SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '최종보고일시"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
'                If gComment_All <> "" Or gComment_Code <> "" Then
'                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & chrCR & gComment_Code & "' "
'                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                               SQL = "UPDATE SPSLMJBDI "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
                SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3' "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                res = SendQuery(gServer, SQL)
                
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow
        
        
        
        '//// 결과테이블에서 그룹코드를 제외한 결과중 빈값이 있는경우 처방/접수 테이블에 업데이트 안함
        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
        res = db_select_Col(gServer, SQL)
        'Save_Raw_Data res & vbCrLf & SQL
        
        If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
        ExamCnt = gReadBuf(0)
        'gReadBuf(0) = "0"
        
        If ExamCnt = "0" Then                                                         '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외) 업데이트
            
            '///////// 처방테이블
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)
            Save_Raw_Data res & vbCrLf & SQL
            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            '////////// 접수 테이블
            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '검사코드"
            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)
            Save_Raw_Data res & vbCrLf & SQL
            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
            
        ElseIf res = -1 Then                                                             '///// 쿼리 에러인경우
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        Else                                                                             '///// 결과가 미입력일때는 업데이트 안함
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_POCT = 1
    End With
End Function

Function Insert_Data_XE2100(ByVal argSpcRow As Integer) As Integer
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
    Dim ExamCode_Remark     As String
    
    Dim State_GM    As String       '//// 그룹/멀티 코드
    Dim State_cnt   As Integer      '//// 그룹/멀티 코드 쪽 변수
    Dim State_G     As String       '//// 그룹코드
    Dim State_M     As String       '//// 멀티코드
    Dim State_B     As String       '//// 배터리코드
    
    
    Dim Send_State      As String
    Dim SQL_LOCAL As String
    

    With frmInterface
        gComment_All = ""
        Insert_Data_XE2100 = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        
        State_GM = ""
        State_cnt = 0
        State_G = ""
        State_M = ""
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

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        
        
        
        '/-------------------------------리마크 처리 때문에 인터페이스에 저장된 코드로 검체를 조회해서 리마크 표시해줄것을 찾음(필요한장비만 열기)
'        SQL = "SELECT EXMN_CD "
'        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
'        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
'        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
'        res = db_select_Col(gServer, SQL)
'
'        j = 0
'        Do While gReadBuf(j) <> ""
'            If ExamCode_Remark <> "" Then
'                ExamCode_Remark = ExamCode_Remark & ",'" & gReadBuf(j) & "'"
'            Else
'                ExamCode_Remark = "'" & gReadBuf(j) & "'"
'            End If
'            j = j + 1
'        Loop
'        If ExamCode_Remark = "" Then ExamCode_Remark = "''"
'
'        For i = 1 To frmInterface.vasTemp.DataRowCnt
'            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
'        Next i
        '/--------------------------------------------------------------------------------------------------------------
        
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                gComment_Code = ""
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
                
                '/----------------------------- 자동리마크 처리 (필요한장비만 열자)
                'Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
                '/-----------------------------
                
                               SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                
                If Mid(Trim(GetText(.vasTemp, iRow, 2)), 1, 2) = "L8" Then
                    Send_State = "1"
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
                Else
                    SQL_LOCAL = ""
                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "SELECT COUNT(EXAMCODE) FROM PAT_RES "
                    SQL_LOCAL = SQL_LOCAL & vbCrLf & " WHERE (REFFLAG <> '' OR PANICFLAG <> '' OR  DELTAFLAG <> '' ) "
                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND panicflag = 'P' "
                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND deltaflag = 'D' "
                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND BARCODE = '" & Trim(lsID) & "' "
                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                    res = db_select_Col(gLocal, SQL_LOCAL)
                    
                    '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                    If CCur(gReadBuf(0)) > 0 Then
                        Send_State = "2"
                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                        SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
                    ElseIf CCur(gReadBuf(0)) = 0 Then
                        Send_State = "3"
                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                        SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
                    End If
                End If
                '/----------------------------- 자동리마크 처리 (필요한장비만 열자)
'                If gComment_All <> "" Or gComment_Code <> "" Then
'                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & chrCR & gComment_Code & "' "
'                End If
                '/-----------------------------
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 2)))
                
                State_cnt = InStr(1, State_GM, "/")
                State_G = Mid(State_GM, 1, State_cnt - 1)
                State_M = Mid(State_GM, State_cnt + 1)
                State_GM = Mid(State_GM, State_cnt + 1)
                State_cnt = InStr(1, State_GM, "/")
                State_B = Mid(State_GM, State_cnt + 1)
                    
                    
                '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
                If Trim(State_G) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
                '/------------------------------------
                
                '/------------------------------------ 결과테이블 멀티코드 상태 업데이트
                If Trim(State_M) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_M) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
            '/------------------------------------
            '/------------------------------------ 결과테이블 배터리코드 상태 업데이트
                If Trim(State_B) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
                        If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf Send_State = "3" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_B) & "' "                                        '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                    
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
            '/------------------------------------
            
            
            '/------------------------------------ 접수테이블 STATE 업데이트
                '////////// 접수 테이블
                SQL = "UPDATE SPSLMJBDI "
                If Send_State = "1" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf Send_State = "2" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf Send_State = "3" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "','" & Trim(State_B) & "') "                    '검사코드"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
        
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If

            '/------------------------------------
            End If
        Next iRow
        
        If Send_State = "" Then: cn_Ser.RollbackTrans: Exit Function
        
        '/------------------------------------ 처방테이블 STATE 업데이트
        '///////// 처방테이블
        SQL = "UPDATE SPSLMJBBI "
        If Send_State = "1" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
        ElseIf Send_State = "2" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        ElseIf Send_State = "3" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        End If
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
        res = SendQuery(gServer, SQL)

        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        '/------------------------------------
        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_XE2100 = 1
    End With
End Function

Function Insert_Data_XE2100_R(ByVal argSpcRow As Integer) As Integer
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
    Dim ExamCode_Remark     As String
    
    
    Dim Send_State      As String
    Dim SQL_LOCAL As String
    

    With frmInterface
        gComment_All = ""
        Insert_Data_XE2100_R = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        lsID = ""
        lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))

        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
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

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        
        SQL = "SELECT EXMN_CD "
        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
        res = db_select_Col(gServer, SQL)
        
        j = 0
        Do While gReadBuf(j) <> ""
            If ExamCode_Remark <> "" Then
                ExamCode_Remark = ExamCode_Remark & ",'" & gReadBuf(j) & "'"
            Else
                ExamCode_Remark = "'" & gReadBuf(j) & "'"
            End If
            j = j + 1
        Loop
        If ExamCode_Remark = "" Then ExamCode_Remark = "''"
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt
            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
        Next i
        

        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                gComment_Code = ""
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
                
                
                Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
                
                
                               SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '중간보고자"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '중간보고일시"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '최종보고자"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '최종보고일시"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                
                If Mid(Trim(GetText(.vasTemp, iRow, 2)), 1, 2) = "L8" Then
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
                Else
                    SQL_LOCAL = ""
                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "Select count(examcode) FROM PAT_RES "
                    SQL_LOCAL = SQL_LOCAL & vbCrLf & " WHERE refflag IS NOT NULL "
                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND panicflag IS NOT NULL "
                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND deltaflag IS NOT NULL "
                    SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' "
                    'SQL_LOCAL = SQL_LOCAL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                    res = db_select_Col(gLocal, SQL_LOCAL)
                    
                    If CCur(gReadBuf(0)) = 0 Then
                        Send_State = "3"
                        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
                    Else
                        Send_State = "1"
                        SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
                    End If
                End If
                
                If gComment_All <> "" Or gComment_Code <> "" Then
                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & chrCR & gComment_Code & "' "
                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLMJBDI "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
                
                If Send_State = "3" Then
                    SQL = "UPDATE SPSLHRRST "
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "                     '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                    SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                    res = SendQuery(gServer, SQL)
                Else
                    SQL = "UPDATE SPSLHRRST "
                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "                     '검사코드"
                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                    SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                    res = SendQuery(gServer, SQL)
                    
                End If
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow
        
        
        
        '//// 결과테이블에서 그룹코드를 제외한 결과중 빈값이 있는경우 처방/접수 테이블에 업데이트 안함
        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        
        ExamCnt = gReadBuf(0)
        gReadBuf(0) = "0"
        
        If ExamCnt = "0" Then                                                         '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외) 업데이트
            
            '///////// 처방테이블
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            '////////// 접수 테이블
            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '검사코드"
            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
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
        Else                                                                             '///// 결과가 미입력일때는 업데이트 안함
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_XE2100_R = 1
    End With
End Function


Function Insert_Data_PhD(ByVal argSpcRow As Integer) As Integer
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

    With frmInterface
        Insert_Data_PhD = -1

        lsID = ""
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))

        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
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

        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        'db_BeginTran gServer
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            If sResult1 = "" Then
                sResult1 = Trim(GetText(.vasTemp, iRow, 3))
            ElseIf sResult1 <> "" And sResult2 = "" Then
                sResult2 = Trim(GetText(.vasTemp, iRow, 3))
                
                If IsNumeric(sResult1) = True Then
                    sResult = sResult2 & "(" & sResult1 & ")"
                ElseIf IsNumeric(sResult2) = True Then
                    sResult = sResult1 & "(" & sResult2 & ")"
                End If
                
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
    
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult & "', "                                           '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult & "', "                                           '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DLTA_YN = '', "                                                            'Delta 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '', "                                                            'Panic 체크"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "', "                                     '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '중간보고자"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '중간보고일시"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '최종보고자"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '최종보고일시"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "                                                        '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                    
            End If

        Next iRow

        '//// 결과테이블에서 그룹코드를 제외한 결과중 빈값이 있는경우 처방/접수 테이블에 업데이트 안함
        SQL = "SELECT EXMN_CD FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
        SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
        res = db_select_Vas(gServer, SQL, .vasTemp1)

        If res = 0 Then                                                                 '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외)
            
            '/////
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If

            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
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
                    
            '/////
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If

            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_PhD = 1
    End With
End Function

Function DELETE_LOCAL_ONE(asBarcode As String, asExamdate As String)
    
    SQL = ""
    SQL = SQL & vbCrLf & "DELETE FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE EXAMDATE = '" & asExamdate & "' "
    SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
    SQL = SQL & vbCrLf & "   AND BARCODE = '" & Trim(asBarcode) & "' "
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function Insert_Data_R_PhD(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
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

    Insert_Data_R_PhD = -1
    With frmInterface
        lsID = ""
        lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))
        
        
        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        
        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
        
        gHIVPosFlag = -1
        

        
        'db_BeginTran gServer
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            If sResult1 = "" Then
                sResult1 = Trim(GetText(.vasTemp, iRow, 3))
            ElseIf sResult1 <> "" And sResult2 = "" Then
                sResult2 = Trim(GetText(.vasTemp, iRow, 3))
                
                If IsNumeric(sResult1) = True Then
                    sResult = sResult2 & "(" & sResult1 & ")"
                ElseIf IsNumeric(sResult2) = True Then
                    sResult = sResult1 & "(" & sResult2 & ")"
                End If
                
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
    
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult & "', "                                           '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult & "', "                                           '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DLTA_YN = '', "                                                            'Delta 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '', "                                                            'Panic 체크"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '중간보고자"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '중간보고일시"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '최종보고자"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '최종보고일시"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                    
            End If

        Next iRow
        
        
        
        
        SQL = "SELECT EXMN_CD FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
        SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NULL "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        
        If res = 0 Then                                                                 '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외)
            SQL = "Update SPSLMJBBI"
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT < 2 "
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        
            SQL = "Update SPSLMJBDI"
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
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
        Else
            SQL = "Update SPSLMJBBI"
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT > 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        
            SQL = "Update SPSLMJBDI"
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT > 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        End If
        
        SQL = ""
    
           
        db_Commit gServer
        Insert_Data_R_PhD = 1
    End With
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
    
    Dim QCCnt           As Integer
    
    With frmInterface
        Insert_Data_QC = -1
        ExamCode_Spec = ""
        lsID = ""
        sCnt = "A"
        QCCnt = 0
        If IsNumeric(Trim(GetText(.vasID, argSpcRow, colBarcode))) = False Then
            lsID = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        Else
            lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        End If
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        
        lsQC_Date = Format(GetDateFull, "yyyymmdd")

        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, RESDATE, EXAMDATE, PID " & vbCrLf & _
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


        
        sResult1 = ""
        sResult2 = ""
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            
            If Trim(GetText(.vasTemp, iRow, 1)) = "TIBC" Then
                sResult1 = Trim(Format(GetText(.vasTemp, iRow, 4), "###0"))
                sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            Else
                sResult1 = Trim(GetText(.vasTemp, iRow, 4))
                sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            End If
            
            If Mid(sResult1, 1, 3) = "-99" Then: sResult1 = ""
            
            
                If Trim(GetText(.vasTemp, iRow, 1)) = "IFCC" Or Trim(GetText(.vasTemp, iRow, 1)) = "eAg" Then
                
                Else
                    If sResult1 <> "" Then
                    
                        If sCnt = "A" Then
                            SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                            SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                            SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                            SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                            'SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                            SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(lsQC_Date) & "' "
                            SQL = SQL & vbCrLf & "  AND RSLT_VALU IS NULL "
                            SQL = SQL & vbCrLf & "GROUP BY RSLT_SQNO "
                            res = db_select_Col(gServer, SQL)
                            sCnt = gReadBuf(0)
                        End If
                        
                        If IsNumeric(sCnt) = True Then
                            SQL = "UPDATE SPSLHQRST "
                            SQL = SQL & vbCrLf & "  SET RSLT_VALU = '" & sResult1 & "', "                        '결과(장비결과)
                            SQL = SQL & vbCrLf & "      RSLT_DT = sysdate, "                                     '결과(수정결과)"
                            SQL = SQL & vbCrLf & "      RSLT_RPTR_ID = '" & gEquipCode & "_INF', "                                                           'Delta 체크"
                            SQL = SQL & vbCrLf & "      AMEN_ID = '" & gEquipCode & "_INF', "
                            SQL = SQL & vbCrLf & "      LOT_NO = '" & Trim(GetText(.vasTemp, iRow, 10)) & "', "
                            SQL = SQL & vbCrLf & "      UPDT_DT = sysdate "                                     '결과입력자"
                            SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                            SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                            SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                            SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                            SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(lsQC_Date) & "' "
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
                            If QCCnt = 0 Then
                                SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                                SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                                SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                                SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                                'SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                                SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(lsQC_Date) & "' "
                                res = db_select_Col(gServer, SQL)
        
                                If gReadBuf(0) = "" Then
                                    QCCnt = "1"
                                Else
                                    QCCnt = CLng(gReadBuf(0)) + 1
                                End If
                            End If
                            
                            If Trim(GetText(.vasTemp, iRow, 2)) <> "" Then
                                SQL = ""
                                SQL = SQL & vbCrLf & "INSERT INTO SPSLHQRST(EXMN_DY   ,EQPM_CD ,SBSN_CD ,LVL_CD  "
                                SQL = SQL & vbCrLf & "                     ,RSLT_SQNO ,EXMN_CD ,RSLT_DT ,RSLT_RPTR_ID "
                                SQL = SQL & vbCrLf & "                     ,RSLT_VALU ,SPCM_NO ,DEL_YN "
                                SQL = SQL & vbCrLf & "                     ,REGI_ID   ,RGST_DT ,AMEN_ID ,UPDT_DT, LOT_NO) "
                                SQL = SQL & vbCrLf & "               VALUES('" & Trim(lsQC_Date) & "', '" & Mid(lsID, 3, 3) & "', '" & Mid(lsID, 6, 3) & "', '" & Mid(lsID, 9, 1) & "', "
                                SQL = SQL & vbCrLf & "                      " & QCCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gEquipCode & "_INF', "
                                SQL = SQL & vbCrLf & "                      '" & sResult1 & "', '" & lsID & "', 'N', "
                                SQL = SQL & vbCrLf & "                      '" & gEquipCode & "_INF', sysdate, '" & gEquipCode & "_INF', sysdate , '" & Trim(GetText(.vasTemp, iRow, 10)) & "') "
                                res = SendQuery(gServer, SQL)
                                
                                If res = -1 Then
                                    SaveQuery SQL
                                    cn_Ser.RollbackTrans
                                    Exit Function
                                End If
                                
                            End If
                        End If
                    End If
                        
                End If
            
        Next iRow
        
        cn_Ser.CommitTrans
        Insert_Data_QC = 1
    End With
End Function

Function Insert_Data_ABI7500(ByVal argSpcRow As Integer) As Integer
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
    Dim Remark_Result   As String
    

    With frmInterface
        Insert_Data_ABI7500 = -1
        ExamCode_Spec = ""
        lsID = ""
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
    
        If lsSpecNo = "" Then: Insert_Data_ABI7500 = -1: Exit Function
        
        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
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
            
            If Trim(GetText(.vasTemp, i, 1)) = "HLA-B27" Then: Remark_Result = Trim(GetText(.vasTemp, argSpcRow, 3))

        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        'db_BeginTran gServer
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
    
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '중간보고자"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '중간보고일시"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '최종보고자"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '최종보고일시"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
                
                If Trim(GetText(.vasTemp, iRow, 1)) = "HLA-B51" Then
                    SQL = SQL & vbCrLf & ",       EXMN_PER_OPNN = 'HLA-B27 = " & Remark_Result & "' "                                                          'Remark 입력
                    Remark_Result = ""
                End If
                
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLMJBDI "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
    
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow
        
        
        
        '//// 결과테이블에서 그룹코드를 제외한 결과중 빈값이 있는경우 처방/접수 테이블에 업데이트 안함
        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        
        ExamCnt = gReadBuf(0)
        gReadBuf(0) = "0"
        
        If ExamCnt = "0" Then                                                         '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외) 업데이트
            
            '///////// 처방테이블
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            '////////// 접수 테이블
            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '검사코드"
            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
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
        Else                                                                             '///// 결과가 미입력일때는 업데이트 안함
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_ABI7500 = 1
    End With
End Function

Function Insert_Data_ABI7500_R(ByVal argSpcRow As Integer) As Integer
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
    Dim Remark_Result   As String

    With frmInterface
        Insert_Data_ABI7500_R = -1
        ExamCode_Spec = ""
        lsID = ""
        lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))
    
        If lsSpecNo = "" Then: Insert_Data_ABI7500_R = -1: Exit Function
        
        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
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
            
            If Trim(GetText(.vasTemp, argSpcRow, 1)) = "HLA-B27" Then: Remark_Result = Trim(GetText(.vasTemp, argSpcRow, 3))

        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        'db_BeginTran gServer
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
    
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '중간보고자"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '중간보고일시"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '최종보고자"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '최종보고일시"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태" (1:입력 , 2:중간보고, 3:최종보고)
                
                If Trim(GetText(.vasTemp, iRow, 1)) = "HLA-B51" Then
                    SQL = SQL & vbCrLf & ",       EXMN_PER_OPNN = '" & Remark_Result & "' "                                                          'Remark 입력
                    Remark_Result = ""
                End If
                
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLMJBDI "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
    
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow
        
        
        
        '//// 결과테이블에서 그룹코드를 제외한 결과중 빈값이 있는경우 처방/접수 테이블에 업데이트 안함
        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        
        ExamCnt = gReadBuf(0)
        gReadBuf(0) = "0"
        
        If ExamCnt = "0" Then                                                         '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외) 업데이트
            
            '///////// 처방테이블
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            '////////// 접수 테이블
            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '검사코드"
            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
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
        Else                                                                             '///// 결과가 미입력일때는 업데이트 안함
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_ABI7500_R = 1
    End With
End Function

Function Insert_Data_QC_R(ByVal argSpcRow As Integer) As Integer
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
    
    Dim QCCnt           As Integer
    
    With frmInterface
        Insert_Data_QC_R = -1
        ExamCode_Spec = ""
        lsID = ""
        sCnt = "A"
        QCCnt = 0
        If IsNumeric(Trim(GetText(.vasRID, argSpcRow, colBarcode))) = False Then
            lsID = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        Else
            lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        End If
        lsSpecNo = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))
        
        lsQC_Date = Format(GetDateFull, "yyyymmdd")

        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, RESDATE, EXAMDATE, PID " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpExamDate.Value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
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


        
        sResult1 = ""
        sResult2 = ""
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
    For iRow = 1 To .vasTemp.DataRowCnt
            
            
            If Trim(GetText(.vasTemp, iRow, 1)) = "TIBC" Then
                sResult1 = Trim(Format(GetText(.vasTemp, iRow, 4), "###0"))
                sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            Else
                sResult1 = Trim(GetText(.vasTemp, iRow, 4))
                sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            End If
            
            
            If Mid(sResult1, 1, 3) = "-99" Then: sResult1 = ""
            
            
                If Trim(GetText(.vasTemp, iRow, 1)) = "IFCC" Or Trim(GetText(.vasTemp, iRow, 1)) = "eAg" Then
                
                Else
                    If sResult1 <> "" Then
                    
                        If sCnt = "A" Then
                            SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                            SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                            SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                            SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                            'SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                            SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(lsQC_Date) & "' "
                            SQL = SQL & vbCrLf & "  AND RSLT_VALU IS NULL "
                            SQL = SQL & vbCrLf & "GROUP BY RSLT_SQNO "
                            res = db_select_Col(gServer, SQL)
                            sCnt = gReadBuf(0)
                        End If
                        
                        If IsNumeric(sCnt) = True Then
                            SQL = "UPDATE SPSLHQRST "
                            SQL = SQL & vbCrLf & "  SET RSLT_VALU = '" & sResult1 & "', "                        '결과(장비결과)
                            SQL = SQL & vbCrLf & "      RSLT_DT = sysdate, "                                     '결과(수정결과)"
                            SQL = SQL & vbCrLf & "      RSLT_RPTR_ID = '" & gEquipCode & "_INF', "                                                           'Delta 체크"
                            SQL = SQL & vbCrLf & "      AMEN_ID = '" & gEquipCode & "_INF', "
                            SQL = SQL & vbCrLf & "      LOT_NO = '" & Trim(GetText(.vasTemp, iRow, 10)) & "', "
                            SQL = SQL & vbCrLf & "      UPDT_DT = sysdate "                                     '결과입력자"
                            SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                            SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                            SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                            SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                            SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(lsQC_Date) & "' "
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
                            If QCCnt = 0 Then
                                SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                                SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                                SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                                SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                                'SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                                SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(lsQC_Date) & "' "
                                res = db_select_Col(gServer, SQL)
        
                                If gReadBuf(0) = "" Then
                                    QCCnt = "1"
                                Else
                                    QCCnt = CLng(gReadBuf(0)) + 1
                                End If
                            End If
                            
                            If Trim(GetText(.vasTemp, iRow, 2)) <> "" Then
                                SQL = ""
                                SQL = SQL & vbCrLf & "INSERT INTO SPSLHQRST(EXMN_DY   ,EQPM_CD ,SBSN_CD ,LVL_CD  "
                                SQL = SQL & vbCrLf & "                     ,RSLT_SQNO ,EXMN_CD ,RSLT_DT ,RSLT_RPTR_ID "
                                SQL = SQL & vbCrLf & "                     ,RSLT_VALU ,SPCM_NO ,DEL_YN "
                                SQL = SQL & vbCrLf & "                     ,REGI_ID   ,RGST_DT ,AMEN_ID ,UPDT_DT, LOT_NO) "
                                SQL = SQL & vbCrLf & "               VALUES('" & Trim(lsQC_Date) & "', '" & Mid(lsID, 3, 3) & "', '" & Mid(lsID, 6, 3) & "', '" & Mid(lsID, 9, 1) & "', "
                                SQL = SQL & vbCrLf & "                      " & QCCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gEquipCode & "_INF', "
                                SQL = SQL & vbCrLf & "                      '" & sResult1 & "', '" & lsID & "', 'N', "
                                SQL = SQL & vbCrLf & "                      '" & gEquipCode & "_INF', sysdate, '" & gEquipCode & "_INF', sysdate , '" & Trim(GetText(.vasTemp, iRow, 10)) & "') "
                                res = SendQuery(gServer, SQL)
                                
                                If res = -1 Then
                                    SaveQuery SQL
                                    cn_Ser.RollbackTrans
                                    Exit Function
                                End If
                                
                            End If
                        End If
                    End If
                        
                End If
            
        Next iRow
        
        cn_Ser.CommitTrans
        Insert_Data_QC_R = 1
    End With
End Function

Function Save_ResMemo(ByVal asRow As Long, asMessage As String)
'메시지 저장하기
    Dim sMessage As String
    
    If asMessage = "" Then
        Exit Function
    End If
    
    sMessage = ""
    
'    SQL = "SELECT MESSAGE "
'    SQL = SQL & vbCrLf & " FROM PAT_RESMEMO  "
'    SQL = SQL & vbCrLf & "WHERE EQUIPNO = '" & gEquip & "' "
'    SQL = SQL & vbCrLf & "  AND BARCODE = '" & Trim(GetText(vasID, asRow, colBarcode)) & "' "
'    SQL = SQL & vbCrLf & "  AND EXAMDATE = '" & Format(dtpToday.Text, "yyyymmdd") & "' "
'    res = db_select_Col(gLocal, SQL)
'
'    sMessage = Trim(gReadBuf(0))
    
'    If sMessage = "" Then
        SQL = "INSERT INTO PAT_RESMEMO "
        SQL = SQL & vbCrLf & "     (EXAMDATE, EQUIPNO, BARCODE, MESSAGE )"
        SQL = SQL & vbCrLf & "VALUES('" & Format(frmInterface.dtpToday, "yyyymmdd") & "', "
        SQL = SQL & vbCrLf & "      '" & gEquip & "',"
        SQL = SQL & vbCrLf & "      '" & Trim(GetText(frmInterface.vasID, asRow, colBarcode)) & "', "
        SQL = SQL & vbCrLf & "      '" & asMessage & "') "
'    Else
'        'sMessage = sMessage & vbCrLf & asMessage
'        sMessage = sMessage & ", " & asMessage

'        SQL = " Update pat_resmemo Set " & vbCrLf & _
'              " message = '" & Trim(sMessage) & "' " & vbCrLf & _
'              " Where examdate = '" & Format(Text_Today.Text, "yyyymmdd") & "' " & vbCrLf & _
'              " And equipno = '" & gEquip & "' " & vbCrLf & _
'              " And barcode = '" & Trim(GetText(vasID, asRow, colBarcode)) & "' "
'    End If
    
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function Pat_Info(asBarcode As String) As Integer
    Dim sBarcode As String
    Dim sSpecNo As String

    Pat_Info = -1
    With frmInterface
        '환자정보 가져오기
        If asBarcode = "" Or IsNumeric(asBarcode) = False Then
            Exit Function
        End If
        '바코드번호로 검체번호 불러오기
        
        SQL = "SELECT FN_LABCVTBCNO(" & Trim(asBarcode) & ") FROM DUAL "
        res = db_select_Col(gServer, SQL)
        
        sSpecNo = Trim(gReadBuf(0))
        
        '환자번호, 환자이름, 주민번호, 성별, 나이
        SQL = "SELECT PID, PT_NM, SEX, AGE "
        SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
        SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
        SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
        'SQL = SQL & vbCrLf & "  AND RSLT_STAT < 2 "
        res = db_select_Col(gServer, SQL)
        
        '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
        
        If res = 1 Then
            SetText .vasList, Trim(sSpecNo), 1, colSpecNo
            SetText .vasList, Trim(gReadBuf(0)), 1, colPID
            SetText .vasList, Trim(gReadBuf(1)), 1, colPName
            SetText .vasList, Trim(gReadBuf(2)), 1, colSex
            SetText .vasList, Trim(gReadBuf(3)), 1, colAge
            
            Pat_Info = 1
        Else
        
            Pat_Info = -1
            SaveQuery (SQL)
        End If
    End With
End Function

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim sBarcode As String
    Dim sSpecNo As String

    Get_Sample_Info = -1
    With frmInterface
        '환자정보 가져오기
        sBarcode = Trim(GetText(.vasID, asRow, colBarcode))   '샘플 바코드 번호
        If sBarcode = "" Or IsNumeric(sBarcode) = False Then
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
        'SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
        'SQL = SQL & vbCrLf & "  AND RSLT_STAT < '2' "
        res = db_select_Col(gServer, SQL)
        
        '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
        
        If res = 1 Then
            SetText .vasID, Trim(sSpecNo), asRow, colSpecNo
            SetText .vasID, Trim(gReadBuf(0)), asRow, colPID
            SetText .vasID, Trim(gReadBuf(1)), asRow, colPName
            SetText .vasID, Trim(gReadBuf(2)), asRow, colSex
            SetText .vasID, Trim(gReadBuf(3)), asRow, colAge
            
            Get_Sample_Info = 1
        Else
        
            Get_Sample_Info = -1
            SaveQuery (SQL)
        End If
    End With
End Function

Function Get_Sample_Info_QC(ByVal asRow As Long) As Integer
    Dim sBarcode As String
    Dim sQCdate  As String
    
    Dim sQCBarcode As String
    
    
    Get_Sample_Info_QC = -1
    With frmInterface
        '환자정보 가져오기
        sBarcode = Trim(GetText(.vasID, asRow, colBarcode))   '샘플 바코드 번호
        'Or (Mid(sBarcode, 1, 2) <> "99" Or Mid(sBarcode, 1, 2) <> "QC")
        If Trim(sBarcode) = "" Then
            Exit Function
        End If
        
        sQCdate = Trim(Format(GetDateFull, "yyyymmdd"))
        
        If Mid(sBarcode, 1, 2) = "99" Then
        
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
        
        Else
            If Mid(sBarcode, 1, 2) = "HC" Or Mid(sBarcode, 1, 2) = "LC" Then sBarcode = Mid(sBarcode, 1, 2)
        
                      SQL = "SELECT EQPM_CD, SBSN_CD, LVL_CD   "
        SQL = SQL & vbCrLf & " FROM SPSLMQMST "
        SQL = SQL & vbCrLf & "WHERE   SBSN_ID = '" & sBarcode & "' "
        SQL = SQL & vbCrLf & "GROUP BY EQPM_CD, SBSN_CD, LVL_CD "
        res = db_select_Col(gServer, SQL)
        sQCBarcode = "99" & gReadBuf(0) & gReadBuf(1) & gReadBuf(2) & "1"
        
        SQL = "SELECT SBSN_NO, '정도관리', '', "
        SQL = SQL & vbCrLf & "                 (SELECT MAX(RSLT_SQNO) + 1 FROM SPSLHQRST "
        SQL = SQL & vbCrLf & "                   WHERE EQPM_CD = '" & gReadBuf(0) & "' "
        SQL = SQL & vbCrLf & "                     AND SBSN_CD = '" & gReadBuf(1) & "' "
        SQL = SQL & vbCrLf & "                     AND LVL_CD  = '" & gReadBuf(2) & "' "
        SQL = SQL & vbCrLf & "                     AND EXMN_DY = '" & sQCdate & "' )"
        SQL = SQL & vbCrLf & " FROM SPSLMQMST "
        SQL = SQL & vbCrLf & "WHERE SBSN_ID = '" & sBarcode & "'"

        End If
        res = db_select_Col(gServer, SQL)
        
        '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
        
        If res = 1 Then
            If Mid(sBarcode, 1, 2) = "99" Then
                SetText .vasID, Trim(sBarcode), asRow, colSpecNo
            Else
                SetText .vasID, Trim(sQCBarcode), asRow, colSpecNo
            End If
            
            SetText .vasID, Trim(gReadBuf(0)), asRow, colPID
            SetText .vasID, Trim(gReadBuf(1)), asRow, colPName
            SetText .vasID, Trim(gReadBuf(2)), asRow, colSex
            SetText .vasID, Trim(gReadBuf(3)), asRow, colAge
            
            If Mid(sBarcode, 1, 2) = "99" Then
                SetText .vasList, Trim(sBarcode), asRow, colSpecNo
            Else
                SetText .vasList, Trim(sQCBarcode), asRow, colSpecNo
            End If
            SetText .vasList, Trim(gReadBuf(0)), 1, colPID
            SetText .vasList, Trim(gReadBuf(1)), 1, colPName
            SetText .vasList, Trim(gReadBuf(2)), 1, colSex
            SetText .vasList, Trim(gReadBuf(3)), 1, colAge
            
            Get_Sample_Info_QC = 1
        Else
        
            Get_Sample_Info_QC = -1
            SaveQuery (SQL)
        End If
    End With
End Function

Function Get_Sample_InfoR(ByVal asRow As Long) As Integer
   Dim sBarcode As String
    Dim sSpecNo As String
    With frmInterface
        Get_Sample_InfoR = -1
        '환자정보 가져오기
        sBarcode = Trim(GetText(.vasRID, asRow, colBarcode))   '샘플 바코드 번호
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
            SetText .vasRID, Trim(sSpecNo), asRow, colSpecNo
            SetText .vasRID, Trim(gReadBuf(0)), asRow, colPID
            SetText .vasRID, Trim(gReadBuf(1)), asRow, colPName
            SetText .vasRID, Trim(gReadBuf(2)), asRow, colSex
            SetText .vasRID, Trim(gReadBuf(3)), asRow, colAge
            
            Get_Sample_InfoR = 1
        Else
        
            Get_Sample_InfoR = -1
        End If
    End With
End Function
Function BCNO_TO_SPECNO(asBarcode As String) As String
    Dim adoComm         As ADODB.Command
    Dim rs_CHANGE       As ADODB.Recordset
    Set adoComm = New ADODB.Command
    Set adoComm.ActiveConnection = cn_Ser 'ADOConnection
    
    
    BCNO_TO_SPECNO = ""
        
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT FN_LABCVTBCNO(?) FROM DUAL "
        
    adoComm.CommandType = adCmdText
    adoComm.CommandText = SQL
    'adoComm.Parameters.Append adoComm.CreateParameter("USE_STR_DY", adDate, , , Now)
    'adoComm.Parameters.Append adoComm.CreateParameter("USE_END_DY", adDate, , , Now)
    'adoComm1.Parameters.Append adoComm1.CreateParameter("FN_LABCVTBCNO", adVarChar, , 10, Trim(strExamCode))
    adoComm.Parameters.Append adoComm.CreateParameter("FN_LABCVTBCNO", adVarChar, , 10, Trim(asBarcode))
    Set rs_CHANGE = New ADODB.Recordset
    rs_CHANGE.Open adoComm, , adOpenStatic, adLockBatchOptimistic
    
    BCNO_TO_SPECNO = rs_CHANGE.Fields(0) & ""
    Set adoComm = Nothing
    rs_CHANGE.Close
    
    
End Function


Function EquipExamCode(asEquipCode As String, asBarcode As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim SpecNo As String

Dim sExamCode_arr
Dim sParam_string   As String



    gEquipExamCode = ""
    gExamRange = ""
    EquipExamCode = ""
    sParam_string = ""
    If Trim(asEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(asEquipCode) & "' "
    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp1)
    
    If frmInterface.vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To frmInterface.vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & "," & Trim(GetText(frmInterface.vasTemp1, i, 1)) & ""
        Else
            sExamCode = Trim(GetText(frmInterface.vasTemp1, i, 1))
        End If
    Next i
    
    sExamCode_arr = Split(sExamCode, ",")
    
    SpecNo = BCNO_TO_SPECNO(asBarcode)
    
    For i = 0 To UBound(sExamCode_arr)
        If sParam_string <> "" Then
            sParam_string = sParam_string & ",?"
        Else
            sParam_string = ",?"
        End If
    Next i
    
    If Len(sParam_string) > 1 Then sParam_string = Mid(sParam_string, 2)
    
    Dim adoComm         As ADODB.Command
    Dim rs_CHANGE       As ADODB.Recordset
    Set adoComm = New ADODB.Command
    Set adoComm.ActiveConnection = cn_Ser 'ADOConnection
        
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = ? "
    SQL = SQL & vbCrLf & "  AND A.EXMN_CD IN (" & sParam_string & ") "
    SQL = SQL & vbCrLf & "GROUP BY A.EXMN_CD "
        
    adoComm.CommandType = adCmdText
    adoComm.CommandText = SQL
    adoComm.Parameters.Append adoComm.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(SpecNo))

    For i = 0 To UBound(sExamCode_arr)
        adoComm.Parameters.Append adoComm.CreateParameter("EXMN_CD", adVarChar, , Len(sExamCode_arr(i)), Trim(sExamCode_arr(i)))
        
    Next i
    
    Set rs_CHANGE = New ADODB.Recordset
    rs_CHANGE.Open adoComm, , adOpenStatic, adLockBatchOptimistic
    If rs_CHANGE.EOF = False Then
        gEquipExamCode = rs_CHANGE.Fields(0) & ""
    End If
    Set adoComm = Nothing
    rs_CHANGE.Close
    
  
    If gEquipExamCode <> "" Then
        gEquipExamCode = Trim(gEquipExamCode)
        
        Set adoComm = New ADODB.Command
        Set adoComm.ActiveConnection = cn_Ser 'ADOConnection
        
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT RSLT_SMNO_SIZE  FROM SPSLMFBIF"
        SQL = SQL & vbCrLf & " WHERE EXMN_CD = ? "
        SQL = SQL & vbCrLf & "   AND USE_END_DY > sysdate "
        
        adoComm.CommandType = adCmdText
        adoComm.CommandText = SQL
        
        adoComm.Parameters.Append adoComm.CreateParameter("EXMN_CD", adVarChar, , Len(gEquipExamCode), Trim(gEquipExamCode))

        Set rs_CHANGE = New ADODB.Recordset
        rs_CHANGE.Open adoComm, , adOpenStatic, adLockBatchOptimistic
        If rs_CHANGE.EOF = False Then
            gExamRange = rs_CHANGE.Fields(0) & ""
        End If
        Set adoComm = Nothing
        rs_CHANGE.Close

    End If
    
    
    
End Function

Function EXAMCODE_LIMIT(asExamCode As String, asResult As String) As String
    Dim Limit_Gubun As String
    Dim Low_Value   As String
    Dim High_Value  As String
    Dim rs_LIMIT        As ADODB.Recordset
    
    
    EXAMCODE_LIMIT = ""
    If IsNumeric(asResult) = False Then EXAMCODE_LIMIT = asResult: Exit Function
    
    
    
    If Trim(asExamCode) = "" Then Exit Function
  
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT ALWN_DVSN, ALWN_LOW, ALWN_HIGH  FROM SPSLMFBIF"
    SQL = SQL & vbCrLf & " WHERE EXMN_CD = ? "
    SQL = SQL & vbCrLf & "   AND USE_END_DY > sysdate "
    
    
    Dim adoComm As ADODB.Command
    
    Set adoComm = New ADODB.Command
    Set adoComm.ActiveConnection = cn_Ser 'ADOConnection
    adoComm.CommandType = adCmdText
    
    adoComm.CommandText = SQL
    'adoComm.Parameters.Append adoComm.CreateParameter("USE_STR_DY", adDate, , , Now)
    'adoComm.Parameters.Append adoComm.CreateParameter("USE_END_DY", adDate, , , Now)
    'adoComm1.Parameters.Append adoComm1.CreateParameter("FN_LABCVTBCNO", adVarChar, , 10, Trim(strExamCode))
    adoComm.Parameters.Append adoComm.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(asExamCode))
    Set rs_LIMIT = New ADODB.Recordset
    rs_LIMIT.Open adoComm, , adOpenStatic, adLockBatchOptimistic
    Set adoComm = Nothing
    
    
    
    Limit_Gubun = rs_LIMIT.Fields("ALWN_DVSN") & ""
    Low_Value = rs_LIMIT.Fields("ALWN_LOW") & ""
    High_Value = rs_LIMIT.Fields("ALWN_HIGH") & ""
    
    If Limit_Gubun = "" Then Exit Function
    
    Select Case Limit_Gubun
        Case "1"    '/하한치만
            If IsNumeric(Low_Value) = False Then Exit Function
            
            If CCur(asResult) < CCur(Low_Value) Then
               asResult = "< " & Low_Value
            End If
            
        Case "2"    '/상한치만
            If IsNumeric(High_Value) = False Then Exit Function
            
            If CCur(asResult) > CCur(High_Value) Then
               asResult = "> " & High_Value
            End If
        Case "3"    '/둘다
            If IsNumeric(High_Value) = False Or IsNumeric(High_Value) = False Then Exit Function
            
            If CCur(asResult) < CCur(Low_Value) Then
               asResult = "< " & Low_Value
            ElseIf CCur(asResult) > CCur(High_Value) Then
               asResult = "> " & High_Value
            End If
    End Select
            
    
    EXAMCODE_LIMIT = asResult
    
End Function


Function EquipExamCode_QC(asEquipCode As String, asBarcode As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim SpecNo As String


    EquipExamCode_QC = ""
    
    If Trim(asEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(asEquipCode) & "' "
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
    
    
    SQL = ""
    SQL = "SELECT QC_EXMN_CD "
    SQL = SQL & vbCrLf & " FROM SPSLMQMST "
    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(asBarcode, 3, 3) & "' "
    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(asBarcode, 6, 3) & "' "
    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(asBarcode, 9, 1) & "' "
    SQL = SQL & vbCrLf & "  AND QC_EXMN_CD IN (" & sExamCode & ") "

    res = db_select_Col(gServer, SQL)

  
    If gReadBuf(0) <> "" Then
        gEquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function

Function TLA_LASC_Search(asStartDate As String, asEndDate As String)
    Dim Server_date As String
    Dim buff As String
    Dim StartDate As String
    Dim EndDate As String
    
    buff = "0.7"
    Server_date = Trim(Format(GetDateFull, "yyyy/mm/dd"))
    StartDate = DateDiff("d", Server_date, asStartDate)
    EndDate = DateDiff("d", Server_date, asEndDate)
    
    If InStr(StartDate, "-") > 0 Then: StartDate = CCur(StartDate) * -1
    If InStr(EndDate, "-") > 0 Then: EndDate = CCur(EndDate) * -1
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT '',  C.SPCM_NO, (SELECT FN_LABCVTPRTBCNO(C.SPCM_NO) FROM DUAL), C.SPCM_SQNO, substr(max(B.WORK_NO), -4),C.PID, C.PT_NM, C.SEX, C.AGE "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B, SPSLMJBBI C "
    SQL = SQL & vbCrLf & " WHERE (C.STAT_DVSN = '' OR C.STAT_DVSN IS NULL) "
    SQL = SQL & vbCrLf & "   AND B.RCPN_DT BETWEEN TO_DATE(" & Format(asStartDate, "yyyymmdd") & ", 'yyyymmdd')"
    SQL = SQL & vbCrLf & "                                     AND TO_DATE(" & Format(asEndDate, "yyyymmdd") & ", 'yyyymmdd') + 0.999999 "
    SQL = SQL & vbCrLf & "   AND C.SPCM_NO = A.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND C.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = C.SPCM_STAT "
    SQL = SQL & vbCrLf & "   AND C.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND C.RSLT_STAT = A.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND C.RSLT_STAT = '0' "
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & " GROUP BY C.SPCM_NO, C.PID, C.PT_NM, C.SEX, C.AGE, C.SPCM_SQNO "
  
    res = db_select_Vas(gServer, SQL, frmInterface.vasID, frmInterface.vasID.DataRowCnt + 1)
    
End Function


Function PAT_List_Search(asStartDate As String, asEndDate As String)
    Dim Server_date As String
    
    Dim buff As String
    Dim StartDate As String
    Dim EndDate As String
    
    buff = "0.7"
    Server_date = Trim(Format(GetDateFull, "yyyy/mm/dd"))
    StartDate = DateDiff("d", Server_date, asStartDate)
    EndDate = DateDiff("d", Server_date, asEndDate)
    
    If InStr(StartDate, "-") > 0 Then: StartDate = CCur(StartDate) * -1
    If InStr(EndDate, "-") > 0 Then: EndDate = CCur(EndDate) * -1

    
With frmInterface
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT '',  B.SPCM_NO, (SELECT FN_LABCVTPRTBCNO(B.SPCM_NO) FROM DUAL),'','', C.PID ,C.PT_NM "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B, SPSLMJBBI C "
    
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND B.SPCM_NO = C.SPCM_NO"
    SQL = SQL & vbCrLf & "   AND B.RCPN_DT BETWEEN TO_DATE(" & Format(asStartDate, "yyyymmdd") & ", 'yyyymmdd')"
    SQL = SQL & vbCrLf & "                                     AND TO_DATE(" & Format(asEndDate, "yyyymmdd") & ", 'yyyymmdd') + 0.999999 "
    
'    SQL = SQL & vbCrLf & "   AND B.RGST_DT BETWEEN SYSDATE - " & (CLng(StartDate) + CCur(buff))
'    SQL = SQL & vbCrLf & "                                     AND SYSDATE - " & CLng(EndDate)
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    
    SQL = SQL & vbCrLf & "   AND C.SPCM_STAT = B.SPCM_STAT "
    SQL = SQL & vbCrLf & "   AND C.RSLT_STAT = B.RSLT_STAT "
    
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND C.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND C.RSLT_STAT = '0' "
    SQL = SQL & vbCrLf & " GROUP BY B.SPCM_NO, C.PID, C.PT_NM "
    
    res = db_select_Vas(gServer, SQL, .vasPatList)
    
    
    
    If res = 1 Then
    ElseIf res = -1 Then
        SaveQuery (SQL)
    End If

End With
End Function

Function LASC_Start_Server(ByVal argSpcRow As Integer) As Integer

'S000000009638527410     ********111001100000kim          gim          guim         000****************************************    <--- 내가 한거
'S00000000     1117559341********110000000000000000000000000000000000000000000000000000****************************************    <--- 지금 하고 있는거
With frmInterface
    gEXAM_CBC = "N"
    gEXAM_Diff = "N"
    gEXAM_Reti = "N"
    gEXAM_CBC_Diff = "N"
    
    Call ClearSpread(.vasTemp1)
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam_CBC & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "
    res = db_select_Vas(gServer, SQL, .vasTemp1)
    
    If res > 0 Then: gEXAM_CBC = "Y"
    
    Call ClearSpread(.vasTemp1)
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam_Diff & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "
    res = db_select_Vas(gServer, SQL, .vasTemp1)
    
    If res > 0 Then: gEXAM_Diff = "Y"
    
    Call ClearSpread(.vasTemp1)
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam_Reti & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "
    res = db_select_Vas(gServer, SQL, .vasTemp1)
    
     If res > 0 Then: gEXAM_Reti = "Y"
    
    Call ClearSpread(.vasTemp1)
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam_CBC_Diff & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "
    res = db_select_Vas(gServer, SQL, .vasTemp1)
    
    If res > 0 Then: gEXAM_CBC_Diff = "Y"
    
    Call Lasc_Order_Make(Trim(GetText(.vasID, argSpcRow, colBarcode)), argSpcRow)
    
End With
End Function

Function Lasc_Order_Make(asBarcode As String, asRow As Integer)
'S000000009638527410     ********111001100000kim          gim          guim         000****************************************    <--- 내가 한거
'S00000000     1117559341********110000000000000000000000000000000000000000000000000000****************************************    <--- 지금 하고 있는거
    Dim Order_Total         As String
    Dim Order_Head          As String
    Dim Order_Barcode       As String
    Dim Order_Order         As String
    Dim Order_OrderCBC      As String
    Dim Order_OrderDiff     As String
    Dim Order_OrderReti     As String
    Dim Order_OrderCBCDiff  As String
    Dim Order_Etc1 As String
    Dim Order_Etc2 As String
    
    '///// 변수 초기화
    Order_Total = ""
    Order_Head = ""
    Order_Barcode = ""
    Order_OrderCBC = "0"
    Order_OrderDiff = "0"
    Order_OrderReti = "0"
    Order_OrderCBCDiff = "0"
    Order_Etc1 = ""
    
    
    Order_Head = "S00000000"
    Order_Barcode = SetSpace(asBarcode, 15, 1)
    Order_Etc1 = "********"
    If gEXAM_CBC = "Y" Then: Order_OrderCBC = "1"
    If gEXAM_Diff = "Y" Then: Order_OrderDiff = "1"
    If gEXAM_Reti = "Y" Then: Order_OrderReti = "1"
    If gEXAM_CBC_Diff = "Y" Then: Order_OrderCBC = "1": Order_OrderDiff = "1"
    Order_Order = Order_OrderCBC & Order_OrderDiff & Order_OrderReti & "0000"
    
    Order_Etc2 = "00000000000000000000000000000000000000000000000****************************************"
    
    
    Order_Total = Order_Head & Order_Barcode & Order_Etc1 & Order_Order & Order_Etc2
    Order_Total = chrSTX & Order_Total & chrETX
    
    SetText frmInterface.vasOrder, Order_Total, frmInterface.vasOrder.DataRowCnt + 1, 1
    SetText frmInterface.vasOrder, CStr(asRow), frmInterface.vasOrder.DataRowCnt, 2

End Function

Function TLA_Start_Server(ByVal argSpcRow As Integer) As Integer
    Dim ExamCount As String
    Dim TLA_Equip As String
    Dim i As Integer
    
    '//////////////검사장비코드 Count
    Dim EQ_DX1 As Integer
    Dim EQ_DX2 As Integer
    Dim EQ_DX3 As Integer
    Dim EQ_DXC As Integer
    Dim EQ_DX0 As Integer
    Dim EQ_D1C As Integer
    Dim EQ_D2C As Integer
    Dim EQ_D3C As Integer
    Dim EQ_D0C As Integer
    Dim EQ_CEN As Integer
    Dim EQ_IML As Integer
    Dim EQ_ELE As Integer
    Dim EQ_SER As Integer
    Dim EQ_COB As Integer
    Dim EQ_VST As Integer
    
    '////////////분주 장비판별
    Dim EQ_NO As String
    Dim EQ_NO1 As String
    Dim EQ_NO2 As String
    Dim EQ_NO3 As String
    Dim EQ_NO_JA As String
    
    '/////////// TLA 모검체 장비명
    Dim TLA_MO As String
    '/////////// TLA 자검체 장비명
    Dim TLA_JA(0 To 6) As String
    '/////////// L8 이면 WorkNo
    Dim A_W_No As String
    '/////////// 채혈 일시
    Dim lsRCPN_DT As String
    
With frmInterface
    TLA_Start_Server = -1
    
    EQ_DX1 = 0
    EQ_DX2 = 0
    EQ_DX3 = 0
    EQ_DXC = 0
    EQ_DX0 = 0
    EQ_D1C = 0
    EQ_D2C = 0
    EQ_D3C = 0
    EQ_D0C = 0
    EQ_CEN = 0
    EQ_IML = 0
    EQ_ELE = 0
    EQ_SER = 0
    EQ_COB = 0
    EQ_VST = 0
    A_W_No = ""
    
    Call ClearSpread(.vasTemp1)
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "

    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
    
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT TO_CHAR(B.BLCL_DT,'YYYY/MM/DD'), C.SPCM_SQNO, substr(MAX(B.WORK_NO),-3)   "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B, SPSLMJBBI C "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = C.SPCM_NO  "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND C.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY B.BLCL_DT, C.SPCM_SQNO "
    
    res = db_select_Col(gServer, SQL)
    
    If Trim(Mid(GetText(frmInterface.vasTemp1, 1, 1), 1, 2)) = "L8" Then
        A_W_No = "W" & Format(gReadBuf(2), "000#")
    Else
        A_W_No = "A" & Format(gReadBuf(1), "000#")
    End If
    
    lsRCPN_DT = Trim(gReadBuf(0))
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    For i = 1 To frmInterface.vasTemp1.DataRowCnt
        If TLA_Equip <> "" Then
            TLA_Equip = TLA_Equip & ",'" & Trim(GetText(.vasTemp1, i, 1)) & "'"
        Else
            TLA_Equip = "'" & Trim(GetText(.vasTemp1, i, 1)) & "'"
        End If
    Next i
    
    Call ClearSpread(.vasTemp1)
    
    If TLA_Equip = "" Then: TLA_Equip = "''"
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EQUIPCODE, B.EQUIPNUMBER "
    SQL = SQL & vbCrLf & "  FROM EQUIPEXAM A, DIVISION B "
    SQL = SQL & vbCrLf & " WHERE A.EQUIPCODE = B.EQUIPCODE "
    SQL = SQL & vbCrLf & "   AND A.EXAMCODE IN (" & TLA_Equip & ") "
'''    SQL = SQL & vbCrLf & "   AND B.EQUIPMAIN = 'Y'"
    SQL = SQL & vbCrLf & " GROUP BY A.EQUIPCODE, B.EQUIPNUMBER "
    res = db_select_Vas(gLocal, SQL, .vasTemp1)
    
    Dim cntEquipNum As Integer
    Dim blMainFlag As Boolean
    Dim cntEquipMain As Integer
    Dim blMoFlag As Boolean
    Dim blJaFlag As Boolean
    
    
    blMainFlag = False
    cntEquipNum = 0
    cntEquipMain = 0
    
    For i = 1 To res
        If IsNumeric(Trim(GetText(.vasTemp1, i, 2))) = True Then
            blMainFlag = True
            cntEquipNum = cntEquipNum + CCur(Trim(GetText(.vasTemp1, i, 2)))
            cntEquipMain = cntEquipMain + 1
            
        End If
    Next
    
    If blMainFlag = True Then '메인장비에 검사가 걸리는 경우
        If cntEquipMain = 2 And cntEquipNum = 7 Then
            EQ_NO = cntEquipNum - 1
        
        Else
            EQ_NO = cntEquipNum
        End If
        
    End If
    
    
    For i = 1 To .vasTemp1.DataRowCnt
        If InStr(1, GetText(.vasTemp1, i, 2), "-") > 0 Then
            If EQ_NO1 = "" Then
                EQ_NO1 = GetText(.vasTemp1, i, 2)
            ElseIf EQ_NO2 = "" Then
                EQ_NO2 = GetText(.vasTemp1, i, 2)
            ElseIf EQ_NO3 = "" Then
                EQ_NO3 = GetText(.vasTemp1, i, 2)
            End If
        
        Else
            If IsNumeric(Trim(GetText(.vasTemp1, i, 2))) = False Then
                If EQ_NO1 = "" Then
                    EQ_NO1 = GetText(.vasTemp1, i, 2)
                ElseIf EQ_NO2 = "" Or EQ_NO1 <> GetText(.vasTemp1, i, 2) Then
                    EQ_NO2 = GetText(.vasTemp1, i, 2)
                ElseIf EQ_NO3 = "" Then
                    EQ_NO3 = GetText(.vasTemp1, i, 2)
                End If
            End If
            
        End If
    Next i

    
    
'''    For i = 1 To .vasTemp1.DataRowCnt
'''        If IsNumeric(GetText(.vasTemp1, i, 2)) = True And InStr(1, GetText(.vasTemp1, i, 2), "-") = 0 Then
'''            If EQ_NO = "" Then
'''                EQ_NO = CCur(GetText(.vasTemp1, i, 2))
'''            Else
'''                EQ_NO = CCur(EQ_NO) + CCur(GetText(.vasTemp1, i, 2))
'''            End If
'''
'''        ElseIf IsNumeric(GetText(.vasTemp1, i, 2)) = True And InStr(1, GetText(.vasTemp1, i, 2), "-") > 0 Then
'''            If EQ_NO1 = "" Then
'''                EQ_NO1 = GetText(.vasTemp1, i, 2)
'''            ElseIf EQ_NO2 = "" Then
'''                EQ_NO2 = GetText(.vasTemp1, i, 2)
'''            ElseIf EQ_NO3 = "" Then
'''                EQ_NO3 = GetText(.vasTemp1, i, 2)
'''            End If
'''
'''        ElseIf IsNumeric(GetText(.vasTemp1, i, 2)) = False Then
'''            If EQ_NO1 = "" Then
'''                EQ_NO1 = GetText(.vasTemp1, i, 2)
'''            ElseIf EQ_NO2 = "" Or EQ_NO1 <> GetText(.vasTemp1, i, 2) Then
'''                EQ_NO2 = GetText(.vasTemp1, i, 2)
'''            ElseIf EQ_NO3 = "" Then
'''                EQ_NO3 = GetText(.vasTemp1, i, 2)
'''            End If
'''        End If
'''    Next i
    
    If EQ_NO1 <> "" Then
        EQ_NO_JA = "'" & EQ_NO1 & "'"

        If EQ_NO2 <> "" Then
            EQ_NO_JA = EQ_NO_JA & ", '" & EQ_NO2 & "'"

            If EQ_NO3 <> "" Then
                EQ_NO_JA = EQ_NO_JA & ", '" & EQ_NO3 & "'"
            End If

        End If

    End If
    
    ClearSpread .vasTemp1
    
    If EQ_NO_JA = "" Then: EQ_NO_JA = "''"
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA, EQUIPNUMBER "
    SQL = SQL & vbCrLf & "  FROM DIVISION "
    If EQ_NO <> "" Then
        SQL = SQL & vbCrLf & "   WHERE EQUIPNUMBER = '" & EQ_NO & "' "
    Else
        SQL = SQL & vbCrLf & "   WHERE EQUIPNUMBER IN (" & EQ_NO_JA & ") "
    End If
    SQL = SQL & vbCrLf & " GROUP BY EQUIPCODE_TLA, EQUIPNUMBER "
    res = db_select_Vas(gLocal, SQL, .vasTemp1)
    
    If EQ_NO <> "" Then
        TLA_MO = Trim(GetText(.vasTemp1, 1, 1))
        
    Else
        blMoFlag = False
        blJaFlag = False
        For i = 1 To res
            If Mid(Trim(GetText(.vasTemp1, i, 2)), 1, 1) = "-" Then
                TLA_MO = Trim(GetText(.vasTemp1, i, 1))
                blMoFlag = True
                Exit For
            End If
        Next
        
        If blMoFlag = False Then
            For i = 1 To res
                If IsNumeric(Trim(GetText(.vasTemp1, i, 2))) = False Then
                    TLA_MO = Trim(GetText(.vasTemp1, i, 1))
                    blMoFlag = True
                    blJaFlag = True
                    Exit For
                End If
            Next
        End If
        
    End If
    ClearSpread .vasTemp1
    
    gReadBuf(0) = ""
    
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA , COUNT(EQUIPCODE_TLA) "
    SQL = SQL & vbCrLf & "  FROM DIVISION  "
    If EQ_NO_JA <> "" Then
        SQL = SQL & vbCrLf & "   WHERE EQUIPNUMBER IN (" & EQ_NO_JA & ") "
        SQL = SQL & vbCrLf & "     AND USEYN = 'Y' "
        SQL = SQL & vbCrLf & "     AND EQUIPCODE_TLA <> '" & TLA_MO & "' "
    Else
        SQL = SQL & vbCrLf & "   WHERE EQUIPNUMBER = '없음' "
    End If
        
    SQL = SQL & vbCrLf & " GROUP BY EQUIPCODE_TLA "
    res = db_select_Vas(gLocal, SQL, .vasTemp1)
    
    Dim Code_TLA As String
    '/////// 조회한 코드 개수 확인해서
    If GetText(.vasTemp1, 1, 2) = "" Then
        Code_TLA = "0"
    Else
        Code_TLA = GetText(.vasTemp1, 1, 2)
    End If
    
    For i = 1 To CInt(res)
        TLA_JA(i - 1) = GetText(.vasTemp1, i, 1)
    Next i
    
    If TLA_JA(0) = "" And blMoFlag = True And blJaFlag = True Then: TLA_JA(0) = TLA_MO
''    i = 0
''    Do While gReadBuf(i) <> ""
''
''        TLA_JA(i) = gReadBuf(i)
''        gReadBuf(i) = ""
''
''        i = i + 1
''    Loop
''
''    i = 0
''    Do While TLA_JA(i) <> ""
''        If TLA_JA(i) = TLA_MO And TLA_JA(i + 1) <> "" Then
''            TLA_JA(i) = TLA_JA(i + 1)
''            TLA_JA(i + 1) = ""
''        ElseIf TLA_JA(i) = TLA_MO And TLA_JA(i + 1) = "" Then
''            TLA_JA(i) = TLA_JA(i + 1)
''            TLA_JA(i + 1) = ""
''        End If
''        i = i + 1
''    Loop
    
    If TLA_MO = "" Then
        TLA_MO = TLA_JA(0)
    End If
    
    
    res = TLA_Division(argSpcRow, TLA_MO, TLA_JA(0), TLA_JA(1), TLA_JA(2), A_W_No, lsRCPN_DT)
    If res = -1 Then
        Save_Raw_Data "[TLA] " & GetDateFull & ":   검체 전송 실패"
        Exit Function
    End If
    
    TLA_Start_Server = 1

End With
End Function

Function TLA_Division(argSpcRow As Integer, asMO As String, _
                      asJA1 As String, asJA2 As String, asJA3 As String, _
                      asA_W_No As String, asRCPN_DT As String) As Integer
                      
    TLA_Division = -1
    Dim BarCodeNo As String
    Dim Age_Conv As String
    Dim i As Integer
    
    Dim Signal                      As String
    Dim Signal_Head                 As String
    Dim Signal_Barcode              As String
    Dim Signal_SpecNo               As String
    Dim Signal_Print                As String
    Dim Signal_UseDate              As String
    Dim Signal_Pname                As String
    Dim Signal_Age                  As String
    Dim Signal_Sex                  As String
    Dim Signal_ReceDate             As String
    Dim Signal_ReceNo               As String
    Dim Signal_Info                 As String
    Dim Signal_Mo                   As String
    Dim Signal_Mo_Bunju             As String
    Dim Signal_MO_Place             As String
    Dim Signal_JA_Bansong(0 To 6)   As String
    Dim Signal_JA_Bunju(0 To 6)     As String
    Dim Signal_JA_Rank(0 To 6)      As String
    Dim Signal_JA_EQName(0 To 6)    As String
    Dim JA_CODE(1 To 3)             As String
    Dim JA_VALUE(1 To 3)            As String
    
    
    Signal_Head = ""
    Signal_Barcode = ""
    Signal_SpecNo = ""
    Signal_Print = ""
    Signal_UseDate = ""
    Signal_Pname = ""
    Signal_Age = ""
    Signal_Sex = ""
    Signal_ReceDate = ""
    Signal_ReceNo = ""
    Signal_Info = ""
    Signal_Mo = ""
    Signal_Mo_Bunju = ""
    Signal_MO_Place = ""
    
    For i = 0 To 6
       Signal_JA_Bansong(i) = ""
       Signal_JA_Bunju(i) = ""
       Signal_JA_Rank(i) = ""
       Signal_JA_EQName(i) = ""
    Next i
    
    
    '//////// 초기화
    Signal = ""
    Signal_Head = ""
    Signal_Barcode = ""
    Signal_SpecNo = ""
    Signal_Print = ""
    Signal_UseDate = ""
    Signal_Pname = ""
    Signal_Age = ""
    Signal_Sex = ""
    Signal_ReceDate = ""
    Signal_ReceNo = ""
    Signal_Info = ""
    Signal_Mo = ""
    Signal_Mo_Bunju = ""
    Signal_MO_Place = ""
    
    '/////// 모검체, 자검체 초기화
    
    
    If asMO = "" Then: Exit Function
With frmInterface
    '//////// 장비에 오더 넣기
    BarCodeNo = Trim(GetText(.vasID, argSpcRow, colBarcode))
    
    Signal_Head = "IC"
    Signal_Barcode = SetSpace(Trim(GetText(.vasID, argSpcRow, colBarcode)), 14, 2)
    Signal_SpecNo = SetSpace(Trim(GetText(.vasID, argSpcRow, colBarcode)), 14, 2)
    
    Signal_Print = SetSpace(Format(Mid(asA_W_No, 2), "0000"), 4, 1) & "/"
    Signal_Print = Signal_Print & Trim(GetText(.vasID, argSpcRow, colPID)) & "/"
    Signal_Print = Signal_Print & "    " & "/"                                          '////// 접수파트 조회해야함
    Signal_Print = Signal_Print & Trim(GetText(.vasID, argSpcRow, colSex)) & ""
    Signal_Print = SetSpace(Signal_Print, 32, 2)
    
    Signal_UseDate = SetSpace(asRCPN_DT, 10)
    Signal_Pname = SetSpace_1(Trim(GetText(.vasID, argSpcRow, colPName)), 14, 2)
    
    
    If IsNumeric(Trim(GetText(.vasID, argSpcRow, colAge))) = True Then
        Signal_Age = CStr(Trim(CCur(Format(Date, "yyyy"))) - CCur(Trim(GetText(.vasID, argSpcRow, colAge))) - 1) & "/01"
    Else
        Age_Conv = Mid(Trim(GetText(.vasID, argSpcRow, colAge)), 1, 2)
        If IsNumeric(Age_Conv) = False Then
            Age_Conv = Mid(Trim(GetText(.vasID, argSpcRow, colAge)), 1, 1)
        End If
        
        If Age_Conv > 11 Then
            Age_Conv = 1
        Else
            Age_Conv = 2
        End If
        
        Signal_Age = CStr(Trim(CCur(Format(Date, "yyyy"))) - CCur(Age_Conv - 1)) & "/01"
    End If
    Signal_Age = SetSpace(Signal_Age, 7)
    
    Signal_Sex = Trim(GetText(.vasID, argSpcRow, colSex))
    Signal_ReceDate = Format(.dtpToday.Value, "yyyy/mm/dd")
    Signal_ReceDate = Mid(Signal_ReceDate, 1, 4) & "/" & Mid(Signal_ReceDate, 6, 2) & "/" & Mid(Signal_ReceDate, 9, 2)
    Signal_ReceNo = "    "
    Signal_Info = "        " 'Mid(asA_W_No, 2)
    Signal_Mo = asMO
    
    If asJA1 <> "" And asJA2 = "" And asJA3 = "" Then
        Signal_Mo_Bunju = "1"
    ElseIf asJA1 <> "" And asJA2 <> "" And asJA3 = "" Then
        Signal_Mo_Bunju = "2"
    ElseIf asJA1 <> "" And asJA2 <> "" And asJA3 <> "" Then
        Signal_Mo_Bunju = "3"
    ElseIf asJA1 = "" And asJA2 = "" And asJA3 = "" Then
        Signal_Mo_Bunju = "0"
    End If
    
    
    Signal_MO_Place = asA_W_No
    
    Signal = Signal_Head & Signal_Barcode & Signal_SpecNo & Signal_Print & Signal_UseDate & Signal_Pname & _
             Signal_Age & Signal_Sex & Signal_ReceDate & Signal_ReceNo & Signal_Info & Signal_Mo & Signal_Mo_Bunju & Signal_MO_Place
    
    If asJA1 = "" Then
        
    ElseIf asJA1 <> "" And asJA2 = "" And asJA3 = "" Then
        ClearSpread .vasTemp1
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA, JA_VALUES"
        SQL = SQL & vbCrLf & "  FROM Division "
        SQL = SQL & vbCrLf & " WHERE EQUIPCODE_TLA = '" & asJA1 & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp1)
        
        For i = 1 To res
            If Trim(GetText(.vasTemp1, i, 2)) <> "" And Trim(GetText(.vasTemp1, i, 2)) <> "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = Format(Trim(GetText(.vasTemp1, i, 2)), "0000")
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            End If
        Next i
        
        Signal = Signal & "    " & JA_VALUE(1) & "     " & JA_CODE(1)
        ClearSpread .vasTemp1
        
    ElseIf asJA1 <> "" And asJA2 <> "" And asJA3 = "" Then
        ClearSpread .vasTemp1
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA, JA_VALUES"
        SQL = SQL & vbCrLf & "  FROM Division "
        SQL = SQL & vbCrLf & " WHERE EQUIPCODE_TLA IN ('" & asJA1 & "', '" & asJA2 & "') "
        res = db_select_Vas(gLocal, SQL, .vasTemp1)
        
        For i = 1 To res
            If Trim(GetText(.vasTemp1, i, 2)) <> "" And Trim(GetText(.vasTemp1, i, 2)) <> "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = Format(Trim(GetText(.vasTemp1, i, 2)), "0000")
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            End If
        Next i
        
        Signal = Signal & "    " & JA_VALUE(1) & "     " & JA_CODE(1)
        Signal = Signal & "      " & JA_VALUE(2) & "     " & JA_CODE(2)
        ClearSpread .vasTemp1
        
                
    ElseIf asJA1 <> "" And asJA2 <> "" And asJA3 <> "" Then
        ClearSpread .vasTemp1
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA, JA_VALUES"
        SQL = SQL & vbCrLf & "  FROM Division "
        SQL = SQL & vbCrLf & " WHERE EQUIPCODE_TLA IN ('" & asJA1 & "', '" & asJA2 & "', '" & asJA3 & "') "
        res = db_select_Vas(gLocal, SQL, .vasTemp1)
        
        For i = 1 To res
            If Trim(GetText(.vasTemp1, i, 2)) <> "" And Trim(GetText(.vasTemp1, i, 2)) <> "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = Format(Trim(GetText(.vasTemp1, i, 2)), "0000")
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            End If
        Next i
        
        Signal = Signal & "    " & JA_VALUE(1) & "     " & JA_CODE(1)
        Signal = Signal & "      " & JA_VALUE(2) & "     " & JA_CODE(2)
        Signal = Signal & "      " & JA_VALUE(3) & "     " & JA_CODE(3)
        ClearSpread .vasTemp1

    End If
End With
     
    Dim FilNum
    Dim sFileName
    FilNum = FreeFile
    
    
    If Dir("c:\his\LIS", vbDirectory) <> "LIS" Then
        MkDir ("c:\his" & "\LIS")
    End If
    
    sFileName = BarCodeNo
    
    If Dir("c:\his\LIS\" & sFileName & ".txt", vbDirectory) <> sFileName & ".txt" Then
        Open "c:\his\LIS\" & sFileName & ".txt" For Append As FilNum
        Print #FilNum, Signal
        Close FilNum
    End If
    
'    Open "c:\his\LIS\" & sFileName & ".txt" For Append As FilNum
'    Print #FilNum, Signal
'    Close FilNum
    
    SQL = ""
    
    TLA_Division = 1
    
    SQL = "UPDATE SPSLMJBBI "
    SQL = SQL & vbCrLf & "   SET STAT_DVSN = 'T' "
    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colSpecNo)) & "' "
    SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colPID)) & "' "
    SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND (STAT_DVSN IS NULL OR STAT_DVSN = '') "
    res = SendQuery(gServer, SQL)
    
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function
'-- 해당 환자 검사의 H/L, Delta, Panic 판정하기
Function GetDecision(ByVal argSpcRow As Integer, ByVal strBarNo As String, ByVal strExamCode As String, ByVal strResult As String) As String
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

    If IsNumeric(strResult) = False Then
        Do
            strResult = Mid(strResult, 2)
            If IsNumeric(Mid(strResult, 1, 1)) = True Then
                If InStr(1, strResult, ")") > 0 Then: strResult = Mid(strResult, 1, InStr(1, strResult, ")") - 1)
                Exit Do
            End If
            If Len(strResult) = 0 Then Exit Do
        Loop
    End If

    '-- 환자의 성별
    strSex = Trim(GetText(frmInterface.vasID, argSpcRow, colSex))

    '-- 해당 환자의 참고치,델타,패닉 찾아오기
    '-- osw 추가 begin
    'ADODB.Command 를 이용한 방법입니다.
    '아래 사용법을 참고하세요
    'Binding Variable 처리
    
        Dim adoComm As ADODB.Command
        
        Set adoComm = New ADODB.Command
        Set adoComm.ActiveConnection = cn_Ser 'ADOConnection
        adoComm.CommandType = adCmdText
        
              SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW   "
        SQL = SQL & " FROM SPSLMFBIF       "
        SQL = SQL & " WHERE USE_STR_DY <= SYSDATE "
        SQL = SQL & "   AND USE_END_DY >= SYSDATE "
        SQL = SQL & "   and EXMN_CD = ? "
        
        adoComm.CommandText = SQL
        '//ex1
    '    adoComm.Parameters.Append adoComm.CreateParameter("USE_STR_DY", adDate, , , Now)
    '    adoComm.Parameters.Append adoComm.CreateParameter("USE_END_DY", adDate, , , Now)
    '    adoComm.Parameters.Append adoComm.CreateParameter("EXMN_CD", adVarChar, , 10, "")
    '
    '    'adoComm.Parameters("USE_STR_DY").Value = "SYSDATE"
    '    'adoComm.Parameters("USE_END_DY").Value = "SYSDATE"
    '    adoComm.Parameters("EXMN_CD").Value = Trim(strExamCode)
        
        '//ex2
        'adoComm.Parameters.Append adoComm.CreateParameter("USE_STR_DY", adDate, , , Now)
        'adoComm.Parameters.Append adoComm.CreateParameter("USE_END_DY", adDate, , , Now)
        adoComm.Parameters.Append adoComm.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(strExamCode))
        
        Set rs_DPRef = New ADODB.Recordset
        rs_DPRef.Open adoComm, , adOpenStatic, adLockBatchOptimistic
        Set adoComm = Nothing
    
    Do Until rs_DPRef.EOF
        '-- 성별로 판정결과 비교
        '-- 결과값이 수치일 경우에만 비교한다.
        If IsNumeric(strResult) = True Then
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
        If IsNumeric(strResult) = True Then
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
        SQL = SQL & vbCrLf & "           AND A.SPCM_NO = FN_LABCVTBCNO(?)                                       "
        SQL = SQL & vbCrLf & "         GROUP BY B.PID, B.EXMN_CD ) A, SPSLHRRST B                                                "
        SQL = SQL & vbCrLf & " WHERE A.PID = B.PID                                                                               "
        SQL = SQL & vbCrLf & "   AND A.LAST_RPTG_DT = B.LAST_RPTG_DT                                                             "
        SQL = SQL & vbCrLf & "   AND A.EXMN_CD = B.EXMN_CD                                                                       "
        SQL = SQL & vbCrLf & "   AND A.EXMN_CD = ? "         '검사코드"
        SQL = SQL & vbCrLf & "   AND B.LAST_RPTG_DT BETWEEN (SYSDATE-30) AND SYSDATE                "           '-- 30일 이내
        
        Dim adoComm1 As ADODB.Command
        
        Set adoComm1 = New ADODB.Command
        Set adoComm1.ActiveConnection = cn_Ser 'ADOConnection
        adoComm1.CommandType = adCmdText
        
        adoComm1.CommandText = SQL
        'adoComm.Parameters.Append adoComm.CreateParameter("USE_STR_DY", adDate, , , Now)
        'adoComm.Parameters.Append adoComm.CreateParameter("USE_END_DY", adDate, , , Now)
        adoComm1.Parameters.Append adoComm1.CreateParameter("FN_LABCVTBCNO", adVarChar, , 10, Trim(strBarNo))
        adoComm1.Parameters.Append adoComm1.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(strExamCode))
        Set rs_Delta = New ADODB.Recordset
        rs_Delta.Open adoComm1, , adOpenStatic, adLockBatchOptimistic
        Set adoComm1 = Nothing
        
        'Set rs_Delta = cn_Ser.Execute(SQL)
        Do Until rs_Delta.EOF
            strBefoRslt = rs_Delta.Fields("BEFO_REAL_RSLT")             '이전결과
            strDestRslt = Trim(strResult)  '현재결과
            If IsNumeric(strBefoRslt) = False Then '///////////////////// 이전결과가 문자가 섞였을때
                Do
                    If Len(strBefoRslt) = 0 Then Exit Do
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
                            strDelta = strDelta / CCur(rs_Delta.Fields("DELTA_TERM_DT"))        '기간당 변화비율
                    Case 4:     '4 기간당 변화차 = 변화차 / 기간
                            strDelta = ""
                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
                            strDelta = CDbl(strDelta) / CCur(rs_Delta.Fields("DELTA_TERM_DT"))  '기간당 변화차
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
            If IsNumeric(strDestRslt) And IsNumeric(strBefoRslt) = True Then
                If IsNumeric(rs_DPRef.Fields("DELT_HIGH")) And IsNumeric(rs_DPRef.Fields("DELT_LOW")) Then
                    If (CDbl(strDestRslt) > rs_DPRef.Fields("DELT_HIGH") Or CCur(strDestRslt) < rs_DPRef.Fields("DELT_LOW")) Then
                        strDelta = "D"
                    Else
                        strDelta = " "
                    End If
                Else
                    strPanic = ""
                End If
            End If
            rs_Delta.MoveNext
        Loop

        rs_DPRef.MoveNext
    Loop

    Set rs_DPRef = Nothing
    Set rs_Delta = Nothing

    GetDecision = strHLVal & "/" & strDelta & "/" & strPanic

End Function

''-- 해당 환자 검사의 H/L, Delta, Panic 판정하기
'Function GetDecision(ByVal argSpcRow As Integer, ByVal strBarNo As String, ByVal strExamCode As String, ByVal strResult As String) As String
'    Dim rs_Delta        As ADODB.Recordset
'    Dim rs_DPRef        As ADODB.Recordset
'    Dim strBefoRslt     As String
'    Dim strDestRslt     As String
'    Dim strHLVal        As String
'    Dim strDelta        As String
'    Dim strPanic        As String
'    Dim strSex          As String
'    Dim strHVal         As String
'    Dim strLVal         As String
'
'    If IsNumeric(strResult) = False Then
'        Do
'            strResult = Mid(strResult, 2)
'            If IsNumeric(Mid(strResult, 1, 1)) = True Then
'                If InStr(1, strResult, ")") > 0 Then: strResult = Mid(strResult, 1, InStr(1, strResult, ")") - 1)
'                Exit Do
'            End If
'            If Len(strResult) = 0 Then Exit Do
'        Loop
'    End If
'
'    '-- 환자의 성별
'    strSex = Trim(GetText(frmInterface.vasID, argSpcRow, colSex))
'
'    '-- 해당 환자의 참고치,델타,패닉 찾아오기
'    SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW                 "
'    SQL = SQL & vbCrLf & " FROM SPSLMFBIF   "
'    SQL = SQL & vbCrLf & " WHERE USE_STR_DY <= SYSDATE  "
'    SQL = SQL & vbCrLf & "   AND USE_END_DY >= SYSDATE  "
'    SQL = SQL & vbCrLf & "   and EXMN_CD = '" & Trim(strExamCode) & "' "
'    Set rs_DPRef = cn_Ser.Execute(SQL)
'    Do Until rs_DPRef.EOF
'        '-- 성별로 판정결과 비교
'        '-- 결과값이 수치일 경우에만 비교한다.
'        If IsNumeric(strResult) = True Then
'            strHLVal = ""
'            If strSex = "M" Then
'                If IsNumeric(rs_DPRef.Fields("MALE_HIGH")) Then
'                    If CDbl(strResult) > CDbl(rs_DPRef.Fields("MALE_HIGH")) Then
'                        strHLVal = "H"
'                    Else
'                        strHLVal = " "
'                    End If
'                Else
'                    strHLVal = ""
'                End If
'
'                If IsNumeric(rs_DPRef.Fields("MALE_LOW")) Then
'                    If Trim(strHLVal) = "" Then
'                        If CDbl(strResult) < CDbl(rs_DPRef.Fields("MALE_LOW")) Then
'                            strHLVal = "L"
'                        Else
'                            strHLVal = " "
'                        End If
'                    End If
'                Else
'                    strHLVal = ""
'                End If
'
'            Else
'                If IsNumeric(rs_DPRef.Fields("FEML_HIGH")) Then
'                    If CDbl(strResult) > CDbl(rs_DPRef.Fields("FEML_HIGH")) Then
'                        strHLVal = "H"
'                    Else
'                        strHLVal = " "
'                    End If
'                Else
'                    strHLVal = ""
'                End If
'                If IsNumeric(rs_DPRef.Fields("FEML_LOW")) Then
'                    If Trim(strHLVal) = "" Then
'                        If (CDbl(strResult) < CDbl(rs_DPRef.Fields("FEML_LOW"))) Then
'                            strHLVal = "L"
'                        Else
'                            strHLVal = " "
'                        End If
'                    End If
'                Else
'                    strHLVal = ""
'                End If
'            End If
'        Else
'            strHLVal = ""
'        End If
'
'        '-- Panic 구분
'        '-- 결과값이 수치일 경우에만 비교한다.
'        If IsNumeric(strResult) = True Then
'            strPanic = ""
'            Select Case Trim(rs_DPRef.Fields("PANC_DVSN"))
'                Case 0:     '0 사용안함
'                        strPanic = ""
'                Case 1:     '1 상한만
'                        If IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
'                            If CDbl(strResult) > rs_DPRef.Fields("PANC_HIGH") Then
'                                strPanic = "P"
'                            Else
'                                strPanic = " "
'                            End If
'                        Else
'                            strPanic = ""
'                        End If
'                Case 2:     '2 하한만
'                        If IsNumeric(rs_DPRef.Fields("PANC_LOW")) Then
'                            If CDbl(strResult) < rs_DPRef.Fields("PANC_LOW") Then
'                                strPanic = "P"
'                            Else
'                                strPanic = " "
'                            End If
'                        Else
'                            strPanic = ""
'                        End If
'                Case 3:     '3 모두 사용
'                        If IsNumeric(rs_DPRef.Fields("PANC_LOW")) And IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
'                            If (CDbl(strResult) < rs_DPRef.Fields("PANC_LOW") Or _
'                                CDbl(strResult) > rs_DPRef.Fields("PANC_HIGH")) Then
'                                strPanic = "P"
'                            Else
'                                strPanic = " "
'                            End If
'                        Else
'                            strPanic = ""
'                        End If
'                Case Else:
'                        strPanic = ""
'            End Select
'        Else
'            strPanic = ""
'        End If
'
'
'
'        '** 이전결과 조회 시작
'        '-- 델타값을 계산하기 위한 이전결과 조회 (한달이내 결과값중 최근값만 조회한다.)
'        SQL = ""
'        SQL = SQL & vbCrLf & "SELECT B.SPCM_NO           BEFO_BCNO                                                               "
'        SQL = SQL & vbCrLf & "     , B.EXMN_CD           BEFO_EXMN_CD                                                            "
'        SQL = SQL & vbCrLf & "     , B.REAL_RSLT         BEFO_REAL_RSLT                                                          "
'        SQL = SQL & vbCrLf & "     , B.VIEW_RSLT         BEFO_VIEW_RSLT                                                          "
'        SQL = SQL & vbCrLf & "     , B.LAST_RPTG_DT     BEFO_FINL_DT                                                             "
'        SQL = SQL & vbCrLf & "     , (SYSDATE - B.LAST_RPTG_DT)  DELTA_TERM_DT                                                   "  '오늘부터의 이전결과 기간을 구한다.
'        SQL = SQL & vbCrLf & "     , B.PID               PID                                                                     "
'        SQL = SQL & vbCrLf & "  FROM (SELECT MAX(B.LAST_RPTG_DT) LAST_RPTG_DT                                                    "
'        SQL = SQL & vbCrLf & "             , B.EXMN_CD                                                                           "
'        SQL = SQL & vbCrLf & "             , B.PID                                                                               "
'        SQL = SQL & vbCrLf & "          FROM SPSLHRRST A, SPSLHRRST B                                                            "
'        SQL = SQL & vbCrLf & "         WHERE A.SPCM_NO   <> B.SPCM_NO                                                            "
'        SQL = SQL & vbCrLf & "           AND A.PID        = B.PID                                                                "
'        SQL = SQL & vbCrLf & "           AND A.EXMN_CD    = B.EXMN_CD                                                            "
'        SQL = SQL & vbCrLf & "           AND A.RCPN_DT   >= B.RCPN_DT                                                            "
'        SQL = SQL & vbCrLf & "           AND B.LAST_RPTG_DT IS NOT NULL                                                          "
'        'SQL = SQL & vbCrLf & "           AND A.RSLT_STAT < '3'                                                                   "
'        SQL = SQL & vbCrLf & "           AND A.SPCM_NO = FN_LABCVTBCNO('" & strBarNo & "')                                       "
'        SQL = SQL & vbCrLf & "         GROUP BY B.PID, B.EXMN_CD ) A, SPSLHRRST B                                                "
'        SQL = SQL & vbCrLf & " WHERE A.PID = B.PID                                                                               "
'        SQL = SQL & vbCrLf & "   AND A.LAST_RPTG_DT = B.LAST_RPTG_DT                                                             "
'        SQL = SQL & vbCrLf & "   AND A.EXMN_CD = B.EXMN_CD                                                                       "
'        SQL = SQL & vbCrLf & "   AND A.EXMN_CD = '" & Trim(strExamCode) & "' "         '검사코드"
'        SQL = SQL & vbCrLf & "   AND B.LAST_RPTG_DT BETWEEN (SYSDATE-30) AND SYSDATE                "           '-- 30일 이내
'        Set rs_Delta = cn_Ser.Execute(SQL)
'        Do Until rs_Delta.EOF
'            strBefoRslt = rs_Delta.Fields("BEFO_REAL_RSLT")             '이전결과
'            strDestRslt = Trim(strResult)  '현재결과
'            If IsNumeric(strBefoRslt) = False Then '///////////////////// 이전결과가 문자가 섞였을때
'                Do
'                    If Len(strBefoRslt) = 0 Then Exit Do
'                    strBefoRslt = Mid(strBefoRslt, 2)
'                    If IsNumeric(Mid(strBefoRslt, 1, 1)) = True Then
'                        If InStr(1, strBefoRslt, ")") > 0 Then: strBefoRslt = Mid(strBefoRslt, 1, InStr(1, strBefoRslt, ")") - 1)
'                        Exit Do
'                    End If
'                Loop
'            End If
'
'            '-- Delta 구분  (아래 로직이 맞는지 검증 필요함...必)
'            '-- 결과값이 수치일 경우에만 비교한다.
'            If IsNumeric(strDestRslt) And IsNumeric(strBefoRslt) = True Then
'                strDelta = ""
'                Select Case Trim(rs_DPRef.Fields("DELT_DVSN"))
'                    Case 0:     '0 사용안함
'                            strDelta = ""
'                    Case 1:     '1 변화차 = 현재결과 - 이전결과
'                            strDelta = ""
'                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'                    Case 2:     '2 변화비율 = 변화차 / 이전결과 * 100
'                            strDelta = ""
'                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'                            strDelta = (CDbl(strDelta) / CDbl(strBefoRslt)) * 100               '변화비율
'                    Case 3:     '3 기간당 변화비율 = 변화비율 / 기간
'                            strDelta = ""
'                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'                            strDelta = (CDbl(strDelta) / CDbl(strBefoRslt)) * 100               '변화비율
'                            strDelta = strDelta / CCur(rs_Delta.Fields("DELTA_TERM_DT"))        '기간당 변화비율
'                    Case 4:     '4 기간당 변화차 = 변화차 / 기간
'                            strDelta = ""
'                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'                            strDelta = CDbl(strDelta) / CCur(rs_Delta.Fields("DELTA_TERM_DT"))  '기간당 변화차
'                    Case 5:     '5 절대변화비율 = 변화차 / 이전결과
'                            strDelta = ""
'                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
'                            strDelta = CDbl(strDelta) / CDbl(strBefoRslt)                       '절대변화비율
'                    Case Else:
'                            strDelta = ""
'                End Select
'            Else
'                strDelta = ""
'            End If
'            '-- Delta 판정
'            If IsNumeric(strDestRslt) And IsNumeric(strBefoRslt) = True Then
'                If IsNumeric(rs_DPRef.Fields("DELT_HIGH")) And IsNumeric(rs_DPRef.Fields("DELT_LOW")) Then
'                    If (CDbl(strDestRslt) > rs_DPRef.Fields("DELT_HIGH") Or CCur(strDestRslt) < rs_DPRef.Fields("DELT_LOW")) Then
'                        strDelta = "D"
'                    Else
'                        strDelta = " "
'                    End If
'                Else
'                    strPanic = ""
'                End If
'            End If
'            rs_Delta.MoveNext
'        Loop
'
'        rs_DPRef.MoveNext
'    Loop
'
'    Set rs_DPRef = Nothing
'
'    GetDecision = strHLVal & "/" & strDelta & "/" & strPanic
'
'End Function

Function Make_Remark(asExamCode As String, asSex As String, asResult As String)
'///////////// 코멘트 생성 (검사당)
    Dim Comment_Gubun As Integer
    Dim Comment_MFGubun As String

    Dim Comment_Code As String      '///////// 판별이후
    Dim Comment_CodeH As String
    Dim Comment_CodeL As String

    Dim Comment_RefMH As String
    Dim Comment_RefML As String
    Dim Comment_RefFH As String
    Dim Comment_RefFL As String

    If asSex = "" Then asSex = "M"
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT cmtdest, cmtflag, CMTCODE, cmtcodeSub, cmhigh, cmlow, cFhigh, cFlow "
    SQL = SQL & vbCrLf & "  FROM EQUIPEXAM "
    SQL = SQL & vbCrLf & " WHERE EXAMCODE = '" & asExamCode & "' "
    SQL = SQL & vbCrLf & ""
    res = db_select_Col(gLocal, SQL)
    If gReadBuf(0) = "" Then Exit Function
    
    Comment_Gubun = CInt(gReadBuf(0))
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

    If Comment_Gubun > 0 Then
        Select Case Comment_Gubun '////////// 0:적용안함, 1: 검사전체적용, 2:해당검사적용
            
            Case "1" '/// 전체적용  // 따로 Function  만듬
                

            Case "2" '/// 해당검사적용

                '///// 0:공통, 1:남/여, 2:사용안함
                If Comment_MFGubun = "0" Then
                    
                    If (asResult >= Comment_RefMH) And Comment_RefMH <> "" Then
                        Comment_Code = Comment_CodeH
                    ElseIf (asResult <= Comment_RefML) And Comment_RefML <> "" Then
                        Comment_Code = Comment_CodeL
                    End If
                    
                    SQL = ""
                    SQL = SQL & vbCrLf & "SELECT CNTS "
                    SQL = SQL & vbCrLf & "  FROM SPSLMFRMK "
                    SQL = SQL & vbCrLf & " WHERE OPNN_CD = '" & Comment_Code & "' "
                    'SQL = SQL & vbCrLf & ""
                    res = db_select_Col(gServer, SQL)
                    
                    gComment_Code = gReadBuf(0)
                    
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
                    
                ElseIf Comment_MFGubun = "2" Then
                    
                    SQL = ""
                    SQL = SQL & vbCrLf & "SELECT CNTS "
                    SQL = SQL & vbCrLf & "  FROM SPSLMFRMK "
                    SQL = SQL & vbCrLf & " WHERE OPNN_CD = '" & Comment_CodeH & "' "
                    SQL = SQL & vbCrLf & ""
                    res = db_select_Col(gServer, SQL)
                    
                    gComment_Code = gReadBuf(0)
                    
                End If
            
        End Select

    End If


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
    
    If res < 1 Then Exit Function
    If gReadBuf(0) = "" Then Exit Function
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


Function RsltState_Check(asSpecNo As String, asExamCode As String) As String '/// 결과 형태 : (그룹코드/멀티코드) : 상태가 중간보고 이하일때
    Dim PRSC_CD_G       As String
    Dim EXMN_CD         As String
    Dim PRSC_CD_M       As String
    Dim PRSC_CD_B       As String
    Dim adoComm         As ADODB.Command
    Dim rs_CHANGE       As ADODB.Recordset
    
    Dim sExamCode_arr
    Dim sExamCode As String
    Dim sParam_string   As String
    
    RsltState_Check = ""
    PRSC_CD_G = " "
    PRSC_CD_M = " "
    PRSC_CD_B = " "
    
    sExamCode = Replace(gAllExam, "'", "")
    sExamCode_arr = Split(sExamCode, ",")
    
    For i = 0 To UBound(sExamCode_arr)
        If sParam_string <> "" Then
            sParam_string = sParam_string & ",?"
        Else
            sParam_string = ",?"
        End If
    Next i
    sParam_string = Mid(sParam_string, 2)
    
    
    Set adoComm = New ADODB.Command
    Set adoComm.ActiveConnection = cn_Ser 'ADOConnection
'    SQL = ""
'    SQL = SQL & vbCrLf & "SELECT DISTINCT "
'    SQL = SQL & vbCrLf & "       R1.PRSC_CD "
'    'SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
'    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
'    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
'    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
'    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
'    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
'    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
'    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
'    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD = '" & asExamCode & "' "
'    SQL = SQL & vbCrLf & "   AND R1.PRSC_CD LIKE ('%G%') "
''    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '2' "
'    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
    
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
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = ? "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD = ? "
    SQL = SQL & vbCrLf & "   AND R1.PRSC_CD LIKE ? "
    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
        
    adoComm.CommandType = adCmdText
    adoComm.CommandText = SQL
    adoComm.Parameters.Append adoComm.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(asSpecNo))
    adoComm.Parameters.Append adoComm.CreateParameter("EXMN_CD", adVarChar, , Len(asExamCode), Trim(asExamCode))
    adoComm.Parameters.Append adoComm.CreateParameter("PRSC_CD", adVarChar, , Len("G%"), Trim("G%"))
    
    Set rs_CHANGE = New ADODB.Recordset
    rs_CHANGE.Open adoComm, , adOpenStatic, adLockBatchOptimistic
    If rs_CHANGE.EOF = False Then
        PRSC_CD_G = rs_CHANGE.Fields(0) & ""
    End If
    Set adoComm = Nothing
    rs_CHANGE.Close

    
    Set adoComm = New ADODB.Command
    Set adoComm.ActiveConnection = cn_Ser 'ADOConnection
'    SQL = ""
'    SQL = SQL & vbCrLf & "SELECT DISTINCT "
'    'SQL = SQL & vbCrLf & "       R1.PRSC_CD "
'    SQL = SQL & vbCrLf & "      R1.EXMN_CD "
'    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
'    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
'    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
'    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
'    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
'    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
'    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = ? "
'    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
'    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN (?) "
'    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT = ? "
'    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    'SQL = SQL & vbCrLf & "       R1.PRSC_CD "
    SQL = SQL & vbCrLf & "      R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = ? "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & sParam_string & ") "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN (?) "
    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < ? "
    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
        
    adoComm.CommandType = adCmdText
    adoComm.CommandText = SQL
    adoComm.Parameters.Append adoComm.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(asSpecNo))
    
    For i = 0 To UBound(sExamCode_arr)
        adoComm.Parameters.Append adoComm.CreateParameter("EXMN_CD", adVarChar, , Len(sExamCode_arr(i)), Trim(sExamCode_arr(i)))
    Next i
    adoComm.Parameters.Append adoComm.CreateParameter("PRSC_CD", adVarChar, , Len("%M%"), Trim("%M%"))
    adoComm.Parameters.Append adoComm.CreateParameter("RSLT_STAT", adVarChar, , 1, Trim("2"))
    
    Set rs_CHANGE = New ADODB.Recordset
    rs_CHANGE.Open adoComm, , adOpenStatic, adLockBatchOptimistic
    If rs_CHANGE.EOF = False Then
        PRSC_CD_M = rs_CHANGE.Fields(0) & ""
    End If
    Set adoComm = Nothing
    rs_CHANGE.Close
    
    
        Set adoComm = New ADODB.Command
    Set adoComm.ActiveConnection = cn_Ser 'ADOConnection
'    SQL = ""
'    SQL = SQL & vbCrLf & "SELECT DISTINCT "
'    'SQL = SQL & vbCrLf & "       R1.PRSC_CD "
'    SQL = SQL & vbCrLf & "      R1.EXMN_CD "
'    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
'    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
'    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
'    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
'    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
'    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
'    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
'    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
'    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('M') "
'    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT = '0' "
'    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    'SQL = SQL & vbCrLf & "       R1.PRSC_CD "
    SQL = SQL & vbCrLf & "      R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = ? "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN (?) "
    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < ? "
    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
        
    adoComm.CommandType = adCmdText
    adoComm.CommandText = SQL
    adoComm.Parameters.Append adoComm.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(asSpecNo))
    
    For i = 0 To UBound(sExamCode_arr)
        adoComm.Parameters.Append adoComm.CreateParameter("EXMN_CD", adVarChar, , Len(sExamCode_arr(i)), Trim(sExamCode_arr(i)))
    Next i
    
    adoComm.Parameters.Append adoComm.CreateParameter("PRSC_CD", adVarChar, , Len("%G%"), Trim("%B%"))
    adoComm.Parameters.Append adoComm.CreateParameter("RSLT_STAT", adVarChar, , 1, Trim("2"))
    
    Set rs_CHANGE = New ADODB.Recordset
    rs_CHANGE.Open adoComm, , adOpenStatic, adLockBatchOptimistic
    If rs_CHANGE.EOF = False Then
        PRSC_CD_B = rs_CHANGE.Fields(0) & ""
    End If
    Set adoComm = Nothing
    rs_CHANGE.Close
    
    
    
    

    res = db_select_Col(gServer, SQL)
       
    If gReadBuf(0) <> "" Then: PRSC_CD_M = gReadBuf(0)
    gReadBuf(0) = ""

    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    'SQL = SQL & vbCrLf & "       R1.PRSC_CD "
    SQL = SQL & vbCrLf & "      R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('B') "
    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '0' "
    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
    res = db_select_Col(gServer, SQL)
       
    If gReadBuf(0) <> "" Then: PRSC_CD_B = gReadBuf(0)
    gReadBuf(0) = ""
    
    
    RsltState_Check = PRSC_CD_G & "/" & PRSC_CD_M & "/" & PRSC_CD_B
End Function
