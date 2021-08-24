Attribute VB_Name = "mod서북병원"
Option Explicit

Public Sub ORDER_SEARCH_서북병원(argOrderDate As String, argOrderCode As String)
    gstrQuy = " SELECT "
    gstrQuy = gstrQuy & vbCrLf & "        A.PID AS CHRTNO, "                                                            '/병록번호
    gstrQuy = gstrQuy & vbCrLf & "        DECODE(A.PRSC_OCRR_DVCD, 'I', '입원', 'O', '외래', '기타') AS IO_SECTION, "   '/외래/입원구분(I:입원, O:외래)
    gstrQuy = gstrQuy & vbCrLf & "        C.DEPT_ENGL_ABNM AS DETPCD, "                                                 '/진료과 약칭(영문)
    gstrQuy = gstrQuy & vbCrLf & "        B.PT_NM AS PATNM, "                                                           '/수진자명
    gstrQuy = gstrQuy & vbCrLf & "        B.SEX_CD AS SEX, "                                                            '/성별
    gstrQuy = gstrQuy & vbCrLf & "        B.RESD_NO_1 AS JUMIN1, "                                                      '/주민번호1
    gstrQuy = gstrQuy & vbCrLf & "        B.RESD_NO_2 AS JUMIN2, "                                                      '/주민번호2
    gstrQuy = gstrQuy & vbCrLf & "        fn_PaGetAge(B.RESD_NO_1, B.RESD_NO_2, B.DOBR, A.PRSC_DATE) AS AGE, "          '/HIS 나이계산 함수
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_DATE AS ORDDATE, "                                                     '/처방일자
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_NO AS ORDSEQ, "                                                                  '/처방번호
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_CD AS ORDCD, "                                                         '/처방코드
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_NM AS ORDNM, "                                                         '/처방명
    '''gstrQuy = gstrQuy & vbCrLf & "        A.CNDT_PRSC_UNIQ_NO AS ORDSEQ, "                                              '/실시처방고유번호(유일함)
    gstrQuy = gstrQuy & vbCrLf & "        A.PRSC_CD, "                                                                  '/처방코드
    gstrQuy = gstrQuy & vbCrLf & "        A.DLVR_MATR, "                                                                '/전달사항
    gstrQuy = gstrQuy & vbCrLf & "        A.SUPT_DEPT_DLVR_MATR "                                                       '/지원부서 전달사항
    gstrQuy = gstrQuy & vbCrLf & "   FROM VPRSCINFN A, TPAPTMASTN B, TZDEPTMSTN C "                                     '/VPRSCINFN(처방조회 VIEW), TPAPTMASTN(환자마스터), TZDEPTMSTN(부서마스터)
    gstrQuy = gstrQuy & vbCrLf & "  WHERE A.PID                 = B.PID "
    gstrQuy = gstrQuy & vbCrLf & "    AND A.MDCR_DPMT_CD        = C.DEPT_CD "
    gstrQuy = gstrQuy & vbCrLf & "    AND A.PRSC_DATE           = '" & Replace(argOrderDate, "-", "") & "' "            '/처방일자
    gstrQuy = gstrQuy & vbCrLf & "    AND A.PRSC_VALD_YN        = 'Y' "                                                 '/원처방 살아있는 처방
    gstrQuy = gstrQuy & vbCrLf & "    AND A.CNDT_PRSC_VALD_YN   = 'Y' "                                                 '/실시처방 살아있는 처방
    gstrQuy = gstrQuy & vbCrLf & "    AND A.PRSC_HSTR_CD        = 'O' "                                                 '/처방History 번호
    gstrQuy = gstrQuy & vbCrLf & "    AND A.CNDT_DATE           = '00000000' "
    gstrQuy = gstrQuy & vbCrLf & "    AND A.PRSC_CD            IN (" & argOrderCode & ") "                           '/처방코드
    
    '/※ A.PRSC_NO(처방번호)가 있으나 변화형태가 불특정하여 A.CNDT_PRSC_UNIQ_NO(실시처방고유번호(유일함))로 사용함.
    '/※ A.CNDT_DATE: ACTING 안된 자료(처치일 경우 오더내릴때 날짜가 생성됨/검사일 경우 날짜가 00000000로 설정됨)-검사결과가 완료되지 않은 자료를 찾을 때 이용함.
End Sub

