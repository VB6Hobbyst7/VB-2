Attribute VB_Name = "mod인천의료원1"
Option Explicit

Public Sub ORDER_SEARCH_인천의료원(argOrderDate As String, argOrderCode As String)
    Select Case gtypEQ_INFO.QUERYTYPE
        Case "1" '/3내과
            gstrQuy = "SELECT "
            gstrQuy = gstrQuy & vbCrLf & "       A.PID, "                           '/병록번호
            gstrQuy = gstrQuy & vbCrLf & "       B.PT_NM, "                         '/수진자명
            gstrQuy = gstrQuy & vbCrLf & "       B.SEX_CD, "                        '/SEX
            gstrQuy = gstrQuy & vbCrLf & "       C.PRSC_CD, "                       '/처방코드
            gstrQuy = gstrQuy & vbCrLf & "       D.PRSC_NM, "                       '/처방명
            gstrQuy = gstrQuy & vbCrLf & "       C.MDCR_DPRT_OGCD, "                '/진료과
            gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_DATE, "                     '/처방일자
            gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_SQNO, "                     '/처방순번
            gstrQuy = gstrQuy & vbCrLf & "       E.RCPN_DVCD, "                     '/환자구분('A' : 입원,'O' : 외래 ,'M' : 종건
            gstrQuy = gstrQuy & vbCrLf & "       TRUNC(MONTHS_BETWEEN(SYSDATE,B.DOBR) / 12) AS AGE "
            gstrQuy = gstrQuy & vbCrLf & "  FROM SY_MSSDPTEXMN   A, "               '/환자검사
            gstrQuy = gstrQuy & vbCrLf & "       SY_PCPMPT       B, "               '/환자 M
            gstrQuy = gstrQuy & vbCrLf & "       SY_MEODPRSC     C, "               '/처방
            gstrQuy = gstrQuy & vbCrLf & "       SY_MEZMPRSCMSTR D, "               '/처방MASTER
            gstrQuy = gstrQuy & vbCrLf & "       SY_HOMMRCPN     E "                '/접수
            gstrQuy = gstrQuy & vbCrLf & " WHERE A.RCPN_DT  BETWEEN TO_DATE('" & argOrderDate & "','YYYY-MM-DD') AND TO_DATE('" & argOrderDate & "','YYYY-MM-DD') + 0.9999 "
            gstrQuy = gstrQuy & vbCrLf & "   AND D.PRSC_CD  IN (" & argOrderCode & ") "
            gstrQuy = gstrQuy & vbCrLf & "   AND C.CDIS_YN   = 'N' "
            gstrQuy = gstrQuy & vbCrLf & "   AND B.PID       = A.PID "
            gstrQuy = gstrQuy & vbCrLf & "   AND C.PID       = A.PID "
            gstrQuy = gstrQuy & vbCrLf & "   AND C.PRSC_DATE = A.PRSC_DATE "
            gstrQuy = gstrQuy & vbCrLf & "   AND C.PRSC_SQNO = A.PRSC_SQNO "
            gstrQuy = gstrQuy & vbCrLf & "   AND D.PRSC_CD   = C.PRSC_CD "
            gstrQuy = gstrQuy & vbCrLf & "   AND E.RCPN_NO   = C.RCPN_NO "
    
        Case "3" '/종합검진
            gstrQuy = "SELECT "
            gstrQuy = gstrQuy & vbCrLf & "       A.PID, "                                               '/병록번호
            gstrQuy = gstrQuy & vbCrLf & "       C.PT_NM, "                                             '/수진자명
            gstrQuy = gstrQuy & vbCrLf & "       C.SEX_CD, "                                            '/SEX
            gstrQuy = gstrQuy & vbCrLf & "       'AT1' AS PRSC_CD, "                                    '/처방코드
            gstrQuy = gstrQuy & vbCrLf & "       '종합검진' AS PRSC_NM, "                               '/처방명
            gstrQuy = gstrQuy & vbCrLf & "       'FM' AS MDCR_DPRT_OGCD, "                              '/진료과
            gstrQuy = gstrQuy & vbCrLf & "       A.RCPN_DATE AS PRSC_DATE, "                            '/접수일자 AS 처방일자
            gstrQuy = gstrQuy & vbCrLf & "       A.HLCH_RCPN_NO AS PRSC_SQNO, "                         '/접수번호 AS 처방SEQ
            gstrQuy = gstrQuy & vbCrLf & "       'M' AS RCPN_DVCD, "                                    '/환자구분
            gstrQuy = gstrQuy & vbCrLf & "       TRUNC(MONTHS_BETWEEN(SYSDATE, C.DOBR) / 12) As AGE "   '/연령
            gstrQuy = gstrQuy & vbCrLf & "  FROM SY_MEPDHLCHRCPN A, "
            gstrQuy = gstrQuy & vbCrLf & "       SY_MEPMPKGMSTR B, "
            gstrQuy = gstrQuy & vbCrLf & "       SY_PCPMPT C "
            gstrQuy = gstrQuy & vbCrLf & " WHERE A.PKG_CD        = B.PKG_CD "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.PID           = C.PID "
            gstrQuy = gstrQuy & vbCrLf & "   AND B.MDEX_TYCD     = 'AT1' "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.RCPN_STAT_CD IN ('R','0','1') " '/R.접수, 0.검사진행, 1.결과완료
            gstrQuy = gstrQuy & vbCrLf & "   AND A.RCPN_DATE     = TO_DATE('" & argOrderDate & "','YYYY-MM-DD') "
            gstrQuy = gstrQuy & vbCrLf & " ORDER BY A.RCPN_DATE"
        
        Case Else
            gstrQuy = "SELECT"
            gstrQuy = gstrQuy & vbCrLf & "       A.PID, "
            gstrQuy = gstrQuy & vbCrLf & "       C.PT_NM, "
            gstrQuy = gstrQuy & vbCrLf & "       C.SEX_CD, "
            gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_CD, "                       '/처방코드
            gstrQuy = gstrQuy & vbCrLf & "       B.PRSC_NM, "
            gstrQuy = gstrQuy & vbCrLf & "       A.MDCR_DPRT_OGCD, "
            gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_DATE, "
            gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_SQNO, "
            gstrQuy = gstrQuy & vbCrLf & "       A.RCPN_DVCD, "                     '/환자구분('A' : 입원,'O' : 외래 ,'M' : 종건
            gstrQuy = gstrQuy & vbCrLf & "       TRUNC(MONTHS_BETWEEN(SYSDATE, C.DOBR) / 12) As AGE "
            gstrQuy = gstrQuy & vbCrLf & "  FROM SY_MEODPRSC A, "
            gstrQuy = gstrQuy & vbCrLf & "       SY_MEZMPRSCMSTR B, "
            gstrQuy = gstrQuy & vbCrLf & "       SY_PCPMPT C "
            gstrQuy = gstrQuy & vbCrLf & " WHERE A.PRSC_DATE = TO_DATE('" & argOrderDate & "','YYYY-MM-DD') "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_CD IN (" & argOrderCode & ") "
            gstrQuy = gstrQuy & vbCrLf & "   AND NVL(A.PRSC_DC_YN ,'*') <> 'Y' "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.CDIS_YN   = 'N' "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.PID       = C.PID "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_CD   = B.PRSC_CD "
    End Select
End Sub
