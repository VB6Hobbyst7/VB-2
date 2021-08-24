Attribute VB_Name = "QC_Proc"
Option Explicit

Const colCheckBox = 1
Const colBarCode = 2
Const colPID = 3
Const colPName = 4
Const colRcnt = 5
Const colState = 6
Const colRStart = 6

' 장비코드 검사코드 검사명 수치결과 문자결과 seq
Const colEquipExam = 1
Const colExamCode = 2
Const colExamName = 3
Const colResValue = 4
Const colResult = 5
Const colSeq = 6
Const colResDate = 7
Const colResTime = 8

Public Function Barcode_Gubun(asBarcode As String) As String
    Dim sBarcode As String
    
    Barcode_Gubun = "P"
    
    sBarcode = Trim(asBarcode)
    
    If UCase(Mid(sBarcode, 7, 1)) = "9" Then
        Barcode_Gubun = "Q"
    Else
        Barcode_Gubun = "P"
    End If
    
End Function

Public Function QBarcode(asBarcode As String) As String
    'QBarcode = Replace(QBarcode, "9", "9")
    QBarcode = asBarcode
End Function

Function Get_QC_Info(argSpread As vaSpread, ByVal asRow As Long) As Integer
    Dim sID As String
    
    Dim lsQCLevel As String
    Dim lsLotNo As String
    Dim lsDate As String
    
    '환자정보 가져오기
    sID = Trim(GetText(argSpread, asRow, colBarCode))   '샘플 바코드 번호
    lsDate = Format(Date, "yyyymmdd")
    
    If sID = "" Then
        Exit Function
    End If
    
    '바코드, 병록번호, 환자명, 검체코드, 검체명
    
    SQL = "SELECT QC_SPCM_NO, LOT_NO, QC_CNTL_CD FROM MSLQCRCPT WHERE QC_SPCM_NO = '" & sID & "' "
    res = db_select_Col(gServer, SQL)

    If res = 1 Then
        lsLotNo = Trim(gReadBuf(1))
        lsQCLevel = Trim(gReadBuf(2))
        
        SetText argSpread, lsLotNo, asRow, colPID
        SetText argSpread, lsQCLevel, asRow, colPName
    End If
    
End Function

Function Select_QC_Exam(asBarcode As String, Optional asExamCode As String = "") As String
    Dim strQC As String
    
    Select_QC_Exam = ""
    
    strQC = "SELECT A.EXMN_CD "
    strQC = strQC & vbCrLf & "FROM MSLQCRSLT A, MSLQCRCPT B "
    strQC = strQC & vbCrLf & "Where A.QC_RCPN_SQNO = B.QC_RCPN_SQNO "
    strQC = strQC & vbCrLf & "AND A.QC_SPCM_NO = B.QC_SPCM_NO "
    strQC = strQC & vbCrLf & "AND A.QC_SPCM_NO = '" & Trim(QBarcode(asBarcode)) & "' "
    If asExamCode = "" Then
        strQC = strQC & vbCrLf & "AND A.EXMN_CD IN (" & gAllExam & ") "
    Else
        strQC = strQC & vbCrLf & "AND A.EXMN_CD IN (" & asExamCode & ") "
    End If
    
    strQC = strQC & vbCrLf & "AND B.QC_PRGR_STAT_CD IN ('01', '03', '04') "
    
    Select_QC_Exam = strQC
End Function

Public Function GetQCExamCode_Equip(argCode As String) As String
'검체번호에 존재하는 장비번호 해당하는 검사코드 가져오기

    Dim i As Integer
    Dim sExamCode As String
     
    sExamCode = ""
    GetQCExamCode_Equip = ""
    ClearSpread frmInterface.vaSpread1
    
    If argCode = "" Then
        Exit Function
    End If
    
    sExamCode = ""
    SQL = "Select ExamCode From EquipExam" & vbCrLf & _
          "Where Equipno = '" & gEquip & "'" & vbCrLf & _
          "  And EquipCode = '" & argCode & "' "
    res = db_select_Vas(gLocal, SQL, frmInterface.vaSpread1)
    
    For i = 1 To frmInterface.vaSpread1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(frmInterface.vaSpread1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(frmInterface.vaSpread1, i, 1)) & "'"
        End If
    Next i
     
    GetQCExamCode_Equip = sExamCode
    
End Function

Function Insert_QC_Data(argSpread As vaSpread, ByVal asRow As Integer) As Integer
    
    Dim lsID As String
    Dim i As Integer
    Dim j As Integer
    Dim sEquipCode As String
    Dim sExamCode As String
    Dim sResValue As String
    Dim sResult As String
    Dim sQC_RSLT_SQNO As String
    Dim sQC_RCPN_SQNO As String
    Dim sPRLL_SQNO As String
    Dim sSD_VALU As String
    Dim sMEAN_VALU As String
    Dim sEXMN_EQPM_CD As String
    Dim sQC_CNTL_CD As String
    Dim sLOT_NO As String
    Dim sQC_SPCM_NO As String
    Dim sEXMN_CD As String
    Dim t_SD_CFCN As String
    Dim t_RULE_CD_VALU  As String
    Dim sTransDate As String
    Dim sTransTime As String
    
    Dim QC_TRANS_YN As Boolean
    
    QC_TRANS_YN = False
    Insert_QC_Data = -1
    
    lsID = ""
    lsID = Trim(GetText(argSpread, asRow, colBarCode))
    
    sTransDate = Format(GetDateFull, "yyyymmdd")
    sTransTime = Format(GetDateFull, "hhmmss")
    
    
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    
    SQL = " Select a.equipcode, a.examcode, a.resvalue, a.resvalue, b.resgubun " & vbCrLf & _
          " From pat_res a, equipexam b " & vbCrLf & _
          " Where a.equipno = b.equipno " & vbCrLf & _
          " And a.examcode = b.examcode " & vbCrLf & _
          " And a.equipcode = b.equipcode " & vbCrLf & _
          " And a.equipno = '" & gEquip & "' " & vbCrLf & _
          " And a.barcode = '" & lsID & "' "
          
    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
'''    sCnt = ""
   cn_Ser.BeginTrans
    '서버로 결과값 저장하기
    For i = 1 To frmInterface.vasTemp.DataRowCnt
        
        
        sEquipCode = Trim(GetText(frmInterface.vasTemp, i, 1))
        sExamCode = GetQCExamCode_Equip(sEquipCode)
        sResValue = Trim(GetText(frmInterface.vasTemp, i, 3))
        sResult = Trim(GetText(frmInterface.vasTemp, i, 4))
        
                  
        
        If sExamCode <> "" And sResValue <> "" Then
            
            sQC_RSLT_SQNO = ""
            sQC_RCPN_SQNO = ""
            sPRLL_SQNO = ""
            sSD_VALU = ""
            sMEAN_VALU = ""
            sEXMN_EQPM_CD = ""
            sQC_CNTL_CD = ""
            sLOT_NO = ""
            sQC_SPCM_NO = ""
            sEXMN_CD = ""
            
            SQL = "SELECT "
            SQL = SQL & vbCrLf & "A.QC_RSLT_SQNO , A.QC_RCPN_SQNO, A.PRLL_SQNO, c.SD_VALU, c.MEAN_VALU "
            SQL = SQL & vbCrLf & ", B.EXMN_EQPM_CD, A.QC_CNTL_CD, A.LOT_NO, A.QC_SPCM_NO, A.EXMN_CD "
            SQL = SQL & vbCrLf & "FROM MSLQCRSLT A, MSLQCRCPT B, MSLQCEXMNM C "
            SQL = SQL & vbCrLf & "Where A.QC_SPCM_NO = B.QC_SPCM_NO "
            SQL = SQL & vbCrLf & "AND A.QC_RCPN_SQNO = B.QC_RCPN_SQNO "
            SQL = SQL & vbCrLf & "AND A.EXMN_CD = C.EXMN_CD "
            SQL = SQL & vbCrLf & "AND A.QC_CNTL_CD = C.QC_CNTL_CD "
            SQL = SQL & vbCrLf & "AND A.LOT_NO = C.LOT_NO "
            SQL = SQL & vbCrLf & "AND A.QC_SPCM_NO = '" & lsID & "' "
            SQL = SQL & vbCrLf & "AND A.EXMN_CD IN (" & sExamCode & ") "
            
            res = db_select_Col(gServer, SQL)
            
'            MsgBox res & vbCrLf & SQL
            
            If res > 0 Then
                sQC_RSLT_SQNO = Trim(gReadBuf(0))
                sQC_RCPN_SQNO = Trim(gReadBuf(1))
                sPRLL_SQNO = Trim(gReadBuf(2))
                sSD_VALU = Trim(gReadBuf(3))
                sMEAN_VALU = Trim(gReadBuf(4))
                sEXMN_EQPM_CD = Trim(gReadBuf(5))
                sQC_CNTL_CD = Trim(gReadBuf(6))
                sLOT_NO = Trim(gReadBuf(7))
                sQC_SPCM_NO = Trim(gReadBuf(8))
                sEXMN_CD = Trim(gReadBuf(9))
               
                t_SD_CFCN = ""
                t_RULE_CD_VALU = ""
                
                
'                If IsNumeric(sResValue) = True And IsNumeric(sSD_VALU) = True And IsNumeric(sMEAN_VALU) = True Then
'                    SQL = "select trunc(NVL((TO_NUMBER('" & sResValue & "') - TO_NUMBER('" & sMEAN_VALU & "')) / NULLIF(TO_NUMBER('" & sSD_VALU & "'), 0), 0),0) " & vbCrLf & _
'                          "from dual "
'
'                    SQL = db_select_Col(gServer, SQL)
'                    t_SD_CFCN = Trim(gReadBuf(0))
'
'                End If
'
'                If Trim(t_SD_CFCN) <> "" Then
'    '                SQL = " sp_msl_calcwestiguardrule('" & sQC_RSLT_SQNO & "', '" & sEXMN_EQPM_CD & "', '" & sQC_CNTL_CD & "', " & _
'    '                      "'" & sLOT_NO & "', '" & sQC_SPCM_NO & "', '" & sEXMN_CD & "', '" & t_SD_CFCN & "', " & _
'    '                      "'') "
'    '                res = db_select_Col(gServer, SQL)
'
'                End If
'
'
'                sQC_RSLT_SQNO = Trim(gReadBuf(0))
'                sQC_RCPN_SQNO = Trim(gReadBuf(1))
'                sPRLL_SQNO = Trim(gReadBuf(2))
'                sSD_VALU = Trim(gReadBuf(3))
'                sMEAN_VALU = Trim(gReadBuf(4))
'                sEXMN_EQPM_CD = Trim(gReadBuf(5))
'                sQC_CNTL_CD = Trim(gReadBuf(6))
'                sLOT_NO = Trim(gReadBuf(7))
'                sQC_SPCM_NO = Trim(gReadBuf(8))
'                sEXMN_CD = Trim(gReadBuf(9))
'                t_SD_CFCN = ""
'                t_RULE_CD_VALU = ""
'
    
    
                'QC결과 업데이트
                SQL = "              Update MS.MSLQCRSLT"
                SQL = SQL & vbCrLf & "  SET rslt_valu = '" & sResValue & "',"
                SQL = SQL & vbCrLf & "      sd_cfcn = '" & t_SD_CFCN & "',"
                SQL = SQL & vbCrLf & "      rule_cd_valu = '" & t_RULE_CD_VALU & "',"
                SQL = SQL & vbCrLf & "      --sttt_use_yn = '',"
                SQL = SQL & vbCrLf & "      --mesr_dvcd = '',"
                SQL = SQL & vbCrLf & "      --mesr_cnts = '',"
                SQL = SQL & vbCrLf & "      rslt_inpt_date = to_char(SYSDATE, 'YYYYMMDD'),"
                SQL = SQL & vbCrLf & "      rslt_inpt_time = to_char(SYSDATE, 'HH24MISS'),"
                SQL = SQL & vbCrLf & "      rslt_inps_id = '" & gExamUID & "',"
                SQL = SQL & vbCrLf & "      --rule_clcl_yn = 'Y',"
                SQL = SQL & vbCrLf & "      --rule_clcl_date = to_char(SYSDATE, 'YYYYMMDD'),"
                SQL = SQL & vbCrLf & "      --rule_clcl_time = to_char(SYSDATE, 'HH24MISS'),"
                SQL = SQL & vbCrLf & "      last_updt_usid = '" & gExamUID & "',"
                SQL = SQL & vbCrLf & "      last_uddt = SYSTIMESTAMP"
                SQL = SQL & vbCrLf & "WHERE QC_RSLT_SQNO = '" & sQC_RSLT_SQNO & "'"
                
                res = SendQuery(gServer, SQL)
'                MsgBox res & vbCrLf & SQL
                
                If res = -1 Then
                    Save_Raw_Data "[QueryErr]" & SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            
'            '접수 테이블 상태변경
'            SQL = "Update MS.MSLQCRCPT"
'            SQL = SQL & vbCrLf & "SET   qc_prgr_stat_cd = '04', --결과"
'            SQL = SQL & vbCrLf & "last_updt_usid = '" & gExamUID & "',"
'            SQL = SQL & vbCrLf & "last_uddt = SYSTIMESTAMP"
'            SQL = SQL & vbCrLf & "Where QC_RCPN_SQNO = '" & sQC_RCPN_SQNO & "'"
'            SQL = SQL & vbCrLf & "AND   PRLL_SQNO = '" & sPRLL_SQNO & "'"
'
'            res = SendQuery(gServer, SQL)
'            If res = -1 Then
'                Save_Raw_Data "[QueryErr]" & SQL
'                db_RollBack gServer
'                Exit Function
'            End If
'''            End If

            QC_TRANS_YN = True
            End If
        End If
        'DoSleep 50
    Next i
    
    
    If QC_TRANS_YN = False Then cn_Ser.RollbackTrans: Exit Function
    cn_Ser.CommitTrans
    
    SQL = "update pat_res " & vbCrLf & _
          " set sendflag = '2' " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(argSpread, asRow, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
                    
    Insert_QC_Data = 1
End Function

'''Public Function sp_QC_Select(asBarcode As String) As String
'''    sp_QC_Select = ""
'''
'''    SQL = "     SELECT"
'''    SQL = SQL & vbCrLf & "        'Y' AS OPTN,"
'''    SQL = SQL & vbCrLf & "        C.QC_CNTL_CD,"
'''    SQL = SQL & vbCrLf & "        D.QC_CNTL_NM,"
'''    SQL = SQL & vbCrLf & "        A.LOT_NO,         --Lot No"
'''    SQL = SQL & vbCrLf & "        A.QC_STDY , --QC시작일자"
'''    SQL = SQL & vbCrLf & "        to_char(to_date(A.QC_STDY, 'YYYYMMDD'), 'YYYY-MM-DD') AS QC_STDY_F,"
'''    SQL = SQL & vbCrLf & "        A.QC_SPCM_NO,     --QC 검체번호"
'''    SQL = SQL & vbCrLf & "        A.PRLL_SQNO , --P"
'''    SQL = SQL & vbCrLf & "        A.EXMN_EQPM_CD , --장비코드"
'''    SQL = SQL & vbCrLf & "        E.EXMN_EQPM_NM , --장비명"
'''    SQL = SQL & vbCrLf & "        B.EXMN_CD , --검사코드"
'''    SQL = SQL & vbCrLf & "        F.EXMN_ABNM , --검사명"
'''    SQL = SQL & vbCrLf & "        A.QC_PRGR_STAT_CD , --상태코드"
'''    SQL = SQL & vbCrLf & "        ii.fn_iim_getcmcdnm('MSL0055', A.QC_PRGR_STAT_CD) AS QC_PRGR_STAT_NM,   --상태"
'''    SQL = SQL & vbCrLf & "        to_char(to_date(A.RCPN_DATE, 'YYYYMMDD'), 'YYYY-MM-DD') AS RCPN_DATE,   --접수일자"
'''    SQL = SQL & vbCrLf & "        to_char(to_date(A.RCPN_TIME, 'HH24MISS'), 'HH24:MI:SS') AS RCPN_TIME,   --접수시간"
'''    SQL = SQL & vbCrLf & "        ii.fn_iim_getusernm(A.RCPS_ID) AS RCPN_PRSN,    --접수자"
'''    SQL = SQL & vbCrLf & "        A.QC_RCPN_SQNO , --QC접수일련번호"
'''    SQL = SQL & vbCrLf & "        A.PRLL_SQNO , --Parallel"
'''    SQL = SQL & vbCrLf & "        B.QC_RSLT_SQNO --QC결과일련번호"
'''
'''    SQL = SQL & vbCrLf & "     FROM MS.MSLQCRCPT A"
'''    SQL = SQL & vbCrLf & "     INNER JOIN MS.MSLQCRSLT B ON A.QC_RCPN_SQNO = B.QC_RCPN_SQNO AND A.PRLL_SQNO = B.PRLL_SQNO AND A.QC_STDY = B.QC_STDY"
'''    SQL = SQL & vbCrLf & "     INNER JOIN MS.MSLQCLOTM C ON C.QC_CNTL_CD = B.QC_CNTL_CD AND A.LOT_NO = C.LOT_NO"
'''    SQL = SQL & vbCrLf & "     INNER JOIN MS.MSLQCCTRLM D ON C.QC_CNTL_CD = D.QC_CNTL_CD"
'''    SQL = SQL & vbCrLf & "     INNER JOIN MS.MSLEQPMM E ON A.EXMN_EQPM_CD = E.EXMN_EQPM_CD"
'''    SQL = SQL & vbCrLf & "     INNER JOIN MS.MSLEXMNM F ON B.EXMN_CD = F.EXMN_CD"
'''    SQL = SQL & vbCrLf & "     Where A.QC_RCPN_SQNO = i_QC_RCPN_SQNO"
'''    SQL = SQL & vbCrLf & "     AND C.QC_CNTL_CD = i_QC_CNTL_CD"
'''    SQL = SQL & vbCrLf & "     AND A.LOT_NO = i_LOT_NO"
'''    SQL = SQL & vbCrLf & "     AND A.QC_STDY = i_QC_STDY"
'''    SQL = SQL & vbCrLf & "     AND   A.PRLL_SQNO = i_PRLL_SQNO"
'''    SQL = SQL & vbCrLf & "     ORDER BY F.SCRN_MARK_SEQ;"
'''
'''    res = db_select_Row(gServer, SQL)
'''
'''End Function


'''PROCEDURE sp_msl_selectqcrcptitem(
'''      i_QC_CNTL_CD            IN VARCHAR2,
'''      i_LOT_NO                IN VARCHAR2,
'''      i_QC_STDY               IN VARCHAR2,
'''      i_QC_RCPN_SQNO          IN NUMBER,
'''      i_PRLL_SQNO             IN VARCHAR2,
'''      i_SCRN_ID               IN VARCHAR2,
'''      i_RC1                   OUT CursorType
'''   )
'''   IS
'''   /*
'''
'''   /******************************************************************************
'''   **  File:
'''   **  Name: pg_msl_diagexamqc.sp_msl_selectqcrcptitem
'''   **  Desc: QC접수: QC접수항목 조회
'''   **  This template can be customized:
'''   **
'''   **  Return values:
'''   **
'''   **  Called by:
'''   **
'''   **  Parameters:
'''   **  Input       Output
'''   **     ----------       -----------
'''   **
'''   **  Auth: 이인호(진료지원)
'''   **  Date: 2010.04.28
'''   **
'''   *******************************************************************************
'''   **  Change History
'''   *******************************************************************************
'''   **  Date:  Author:    Description:
'''   **  --------  --------    -------------------------------------------
'''   **
'''   *******************************************************************************/
'''
'''
'''   BEGIN
'''
'''     --이전 접수 내역정보를 이용한 Repeat 조회 등록
'''      If i_QC_RCPN_SQNO Is Not Null Then
'''
'''         OPEN i_RC1 FOR
'''         SELECT
'''            'Y' AS OPTN,
'''            C.QC_CNTL_CD,
'''            D.QC_CNTL_NM,
'''            A.LOT_NO,         --Lot No
'''            A.QC_STDY , --QC시작일자
'''            to_char(to_date(A.QC_STDY, 'YYYYMMDD'), 'YYYY-MM-DD') AS QC_STDY_F,
'''            A.QC_SPCM_NO,     --QC 검체번호
'''            A.PRLL_SQNO , --P
'''            A.EXMN_EQPM_CD , --장비코드
'''            E.EXMN_EQPM_NM , --장비명
'''            B.EXMN_CD , --검사코드
'''            F.EXMN_ABNM , --검사명
'''            A.QC_PRGR_STAT_CD , --상태코드
'''            ii.fn_iim_getcmcdnm('MSL0055', A.QC_PRGR_STAT_CD) AS QC_PRGR_STAT_NM,   --상태
'''            to_char(to_date(A.RCPN_DATE, 'YYYYMMDD'), 'YYYY-MM-DD') AS RCPN_DATE,   --접수일자
'''            to_char(to_date(A.RCPN_TIME, 'HH24MISS'), 'HH24:MI:SS') AS RCPN_TIME,   --접수시간
'''            ii.fn_iim_getusernm(A.RCPS_ID) AS RCPN_PRSN,    --접수자
'''            A.QC_RCPN_SQNO , --QC접수일련번호
'''            A.PRLL_SQNO , --Parallel
'''            B.QC_RSLT_SQNO --QC결과일련번호
'''
'''         FROM MS.MSLQCRCPT A
'''         INNER JOIN MS.MSLQCRSLT B ON A.QC_RCPN_SQNO = B.QC_RCPN_SQNO AND A.PRLL_SQNO = B.PRLL_SQNO AND A.QC_STDY = B.QC_STDY
'''         INNER JOIN MS.MSLQCLOTM C ON C.QC_CNTL_CD = B.QC_CNTL_CD AND A.LOT_NO = C.LOT_NO
'''         INNER JOIN MS.MSLQCCTRLM D ON C.QC_CNTL_CD = D.QC_CNTL_CD
'''         INNER JOIN MS.MSLEQPMM E ON A.EXMN_EQPM_CD = E.EXMN_EQPM_CD
'''         INNER JOIN MS.MSLEXMNM F ON B.EXMN_CD = F.EXMN_CD
'''         Where A.QC_RCPN_SQNO = i_QC_RCPN_SQNO
'''         AND C.QC_CNTL_CD = i_QC_CNTL_CD
'''         AND A.LOT_NO = i_LOT_NO
'''         AND A.QC_STDY = i_QC_STDY
'''         AND   A.PRLL_SQNO = i_PRLL_SQNO
'''         ORDER BY F.SCRN_MARK_SEQ;
'''
'''
'''
'''      --검체번호로 신규 조회 등록
'''      Else
'''
'''         OPEN i_RC1 FOR
'''         SELECT
'''
'''            'Y' AS OPTN,
'''            A.LOT_NO,
'''            A.QC_STDY , --QC시작일자
'''            to_char(to_date(A.QC_STDY, 'YYYYMMDD'), 'YYYY-MM-DD') AS QC_STDY_F,
'''            A.QC_CNTL_CD,
'''            D.QC_CNTL_NM,
'''            '0' AS PRLL_SQNO,
'''            C.EXMN_EQPM_CD,
'''            E.EXMN_EQPM_NM,
'''            B.EXMN_CD,
'''            F.EXMN_ABNM,
'''            '' AS QC_PRGR_STAT_CD,
'''            '' AS QC_PRGR_STAT_NM,
'''            '' AS RCPN_DATE,
'''            '' AS RCPN_TIME,
'''            '' AS RCPN_PRSN,
'''            '' AS QC_RCPN_SQNO,
'''            '' AS PRLL_SQNO,
'''            '' AS QC_RSLT_SQNO
'''
'''         FROM MS.MSLQCLOTM A
'''         INNER JOIN MS.MSLQCEXMNM B ON A.QC_CNTL_CD = B.QC_CNTL_CD AND A.LOT_NO = B.LOT_NO AND A.QC_STDY = B.QC_STDY
'''         INNER JOIN MS.MSLQCCTRLM C ON A.QC_CNTL_CD = C.QC_CNTL_CD
'''         INNER JOIN MS.MSLQCCTRLM D ON C.QC_CNTL_CD = D.QC_CNTL_CD
'''         INNER JOIN MS.MSLEQPMM E ON C.EXMN_EQPM_CD = E.EXMN_EQPM_CD
'''         INNER JOIN MS.MSLEXMNM F ON B.EXMN_CD = F.EXMN_CD
'''         WHERE A.QC_CNTL_CD || A.Lot_No || A.QC_STDY = i_SCRN_ID
'''     --    WHERE A.QC_SPCM_NO = i_QC_SPCM_NO
'''         ORDER BY B.MARK_SEQ;
'''
'''
'''      END IF;
'''
'''   END;

