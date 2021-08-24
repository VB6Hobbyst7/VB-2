VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISLibrarySql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsLibrarySql.cls (���ֿ�������)
'   �ۼ���  : �̻��
'   ��  ��  : IISLibrary���� ����ϴ� ����
'   �ۼ���  : 2004-10-05
'   ��  ��  :
'       1. 1.0.1: �̻��(2004-11-17)
'          - SelectAccInfo �Լ��� ��������
'       2. 1.0.4: �̻��(2004-12-13)
'       3. 1.1.4: �̻��(2004-12-17)
'          - SelectTargetSpcs �Լ��߰�(�˻�����ȸ���)
'       4. 1.1.7: �̻��(2005-04-06)
'       5. 1.1.8: �̻��(2005-04-11)
'       6. 1.2.0: �̻��(2005-05-30)
'       7. 1.2.2: �̻��(2005-06-14)
'   ��  ��  : NOTE Ȯ���Ұ�!
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   ��� : �������̽��� ����ϴ� �������� ��ȸ����
'   �μ� :
'       - pSpcYy : ��ü����
'       - pSpcNo : ��ü��ȣ
'-----------------------------------------------------------------------------'
Public Function SelectAccInfo(ByVal pSpcYy As String, ByVal pSpcNo As String) As String
    Dim SQL As String
    
    '## 1.0.1
    '   - ��ü��, ������ڵ�, �����ڵ�, ä���Ͻ�, �����Ͻø� ���� ��ȸ�ϵ��� ����
    
    '## ��������
    'SQL = " SELECT a.ptid AS PTID, b." & FPTNM & " AS NAME, " & FJUMIN & " AS SSN," & _
          "     a.ageday AS AGEDAY, a.sex AS SEX, c." & FDEPTNM & " AS DEPTNM, d." & FDOCTNM & " AS DOCTNM," & _
          "     a.workarea AS WORKAREA, a.accdt AS ACCDT, a.accseq AS ACCSEQ, a.stscd AS STSCD," & _
          "     a.spccd AS SPCCD, a.reqtotcnt AS REQTOTCNT, a.reqinputcnt AS INPUTCNT, e." & FWARDNM & " AS WARDNM," & _
          "     a.hosilid AS HOSILID, a.testdiv AS TESTDIV, a.qcfg AS QCFG, a.statfg AS STATFG, a.buildcd AS BUILDCD" & _
          " FROM " & TIIS203 & " a, " & THIS001 & " b, " & THIS002 & " c, " & THIS004 & " d, " & THIS003 & " e" & _
          " WHERE a.spcyy='" & pSpcYy & "' AND a.spcno=" & pSpcNo & _
          "     AND a.ptid=b." & FPTID & "(+) AND a.deptcd=c." & FDEPTCD & "(+) AND a.orddoct=d." & FDOCTCD & "(+)" & _
          "     AND a.wardid=e." & FWARDCD & "(+)"

    '## ������ ����
    '## ȯ��ID, �����, ����, ó���ǿ� Outer������ �������� QC��ü������
    '## QC��ü�� ȯ��, �����, ����, ó���ǰ� ����
    SQL = " SELECT DISTINCT a.ptid AS PTID, b.ptnm AS NAME, b.jumin AS SSN, a.ageday AS AGEDAY," & _
          "     a.sex AS SEX, a.workarea AS WORKAREA, a.accdt AS ACCDT, a.accseq AS ACCSEQ," & _
          "     a.stscd AS STSCD, a.spccd AS SPCCD, f.field3 AS SPCNM, a.statfg AS STATFG," & _
          "     a.reqtotcnt AS REQTOTCNT, a.reqinputcnt AS INPUTCNT," & _
          "     a.coldt AS COLDT, a.coltm AS COLTM, a.rcvdt AS RCVDT, a.rcvtm AS RCVTM," & _
          "     a.deptcd AS DEPTCD, '' AS DEPTNM, '' AS DOCTNM," & _
          "     a.wardid AS WARDID, '' AS WARDNM," & _
          "     a.hosilid AS HOSILID, a.testdiv AS TESTDIV, a.qcfg AS QCFG, a.buildcd AS BUILDCD" & _
          " FROM " & TIIS203 & " a, s2ord101 b, " & TIIS002 & " f" & _
          " WHERE a.ptid='" & pSpcYy & "' " & _
          "     AND a.ptid=b." & FPTID & "(+) " & _
          "     AND a.spccd=f.cdval1(+) AND f.cdindex(+)='" & CSPCCD & "'"

    SelectAccInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽��� ����ϴ� ������� ��ȸ����(��ȿ����, D/P��ġ�� ��ȸ)
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'-----------------------------------------------------------------------------'
Public Function SelectResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    '## 1.1.7: �̻��(2005-04-06)
    '   - �������(rsttype)�� LAB001���� ��ȸ�ϵ��� ����
    '## ������ ����� ����
'    SQL = " SELECT a.testcd AS TESTCD, b.abbrnm10 AS TESTNM10, b.testnm AS TESTNM, c.field3 AS SPCNM," & _
'          "     a.rstunit AS UNIT, a.lastrst AS LASTRST, b.rsttype AS RSTTYPE," & _
'          "     a.rstdiv AS RSTDIV, a.detailfg AS DETAILFG, d.avalval AS AVALVAL," & _
'          "     d.panicfg AS PANICFG, d.panicfrval AS PANICFRVAL, d.panictoval AS PANICTOVAL," & _
'          "     d.deltafg AS DELTAFG, d.deltaval AS DELTAFRVAL, d.deltaval2 AS DELTATOVAL" & _
'          " FROM " & TIIS204 & " a, " & TIIS101 & " b, " & TIIS002 & " c, " & TIIS102 & " d" & _
'          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
'          "     AND a.accseq=" & CStr(pAccSeq) & " AND (vfydt IS NULL OR vfydt='')" & _
'          "     AND c.cdindex='" & CSPCCD & "' AND a.testcd=b.testcd" & _
'          "     AND b.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
'          "     AND a.spccd=c.cdval1" & _
'          "     AND a.testcd=d.testcd AND a.spccd=d.spccd" & _
'          "     AND d.applydt=(SELECT MAX(applydt) FROM " & TIIS102 & _
'          "                    WHERE testcd=a.testcd AND spccd=a.spccd)" & _
'          " ORDER BY b.rptseq"

    '## 1.1.7: �̻��(2005-04-06)
    '   - �������(rsttype)�� LAB001���� ��ȸ�ϵ��� ����
    '## 1.0.4: �̻��(2004-12-13)
    '   - QC��ü�� ��ȸ�ȵǴ� ��������!
    '## ���� ����� ����
    SQL = " SELECT a.testcd AS TESTCD, b.abbrnm10 AS TESTNM10, b.testnm AS TESTNM, c.field3 AS SPCNM," & _
          "     a.rstunit AS UNIT, a.lastrst AS LASTRST, b.rsttype AS RSTTYPE," & _
          "     a.rstdiv AS RSTDIV, a.detailfg AS DETAILFG, d.avalval AS AVALVAL," & _
          "     d.panicfg AS PANICFG, d.panicfrval AS PANICFRVAL, d.panictoval AS PANICTOVAL," & _
          "     d.deltafg AS DELTAFG, d.deltaval AS DELTAFRVAL, d.deltaval2 AS DELTATOVAL" & _
          " FROM " & TIIS204 & " a, " & TIIS101 & " b, " & TIIS002 & " c, " & TIIS102 & " d" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND (a.vfydt IS NULL OR a.vfydt='')" & _
          "     AND (a.rstcd IS NULL OR a.rstcd='')" & _
          "     AND c.cdindex='" & CSPCCD & "' AND a.testcd=b.testcd" & _
          "     AND b.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
          "     AND a.spccd=c.cdval1" & _
          "     AND a.testcd=d.testcd AND a.spccd=d.spccd" & _
          "     AND d.applydt=(SELECT MAX(applydt) FROM " & TIIS102 & _
          "                    WHERE testcd=a.testcd AND spccd=a.spccd)" & _
          " ORDER BY b.rptseq"
          
    SelectResultInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��� ���۳��� �Է�����
'   �μ� :
'       - pEqpCd    : ����ڵ�
'       - pTransDt  : ��񿡼� ��� ��������
'       - pTransTm  : ��񿡼� ��� ���ҽð�
'       - pSpcYy    : ��ü����
'       - pSpcNo    : ��ü��ȣ
'       - pTransNo  : ��񿡼� ��ü��ġ
'-----------------------------------------------------------------------------'
Public Function InsertEqpTrans(ByVal pEqpCd As String, ByVal pTransDt As String, _
                    ByVal pTransTm As String, ByVal pSpcYy As String, ByVal pSpcNo As String, _
                    ByVal pTransNo As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS405 & " (eqpcd, transdt, transtm, spcyy, spcno, transno)" & _
          " VALUES ('" & pEqpCd & "', '" & pTransDt & "', '" & pTransTm & "'," & _
          "     '" & pSpcYy & "', " & pSpcNo & ", '" & pTransNo & "')"
    
    InsertEqpTrans = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��� ���۳��� ��������(����ڵ�, ��ü����, ��ü��ȣ�� ��ġ�ϴ°͸� ����)
'   �μ� :
'       - pEqpCd    : ����ڵ�
'       - pSpcYy    : ��ü����
'       - pSpcNo    : ��ü��ȣ
'-----------------------------------------------------------------------------'
Public Function DeleteEqpTrans(ByVal pEqpCd As String, ByVal pSpcYy As String, _
                    ByVal pSpcNo As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TIIS405 & _
          " WHERE eqpcd='" & pEqpCd & "' AND spcyy='" & pSpcYy & "' AND spcno=" & pSpcNo
    
    DeleteEqpTrans = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽����� ����ϴ� �󼼸��ڵ� ����Է� ����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'       - pAutoFg   : ������� ����
'       - pEqpCd    : ����ڵ�
'-----------------------------------------------------------------------------'
Public Function UpdateIISResultDetail(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, _
                    ByVal pAutoFg As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS204 & " SET autofg='" & pAutoFg & "', eqpcd='" & pEqpCd & "', " & _
          " transdt = TO_CHAR(sysdate, 'YYYYMMDD'), transtm = TO_CHAR(sysdate, 'HHMMSS') " & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
          
    UpdateIISResultDetail = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽����� ����ϴ� �󼼸��ڵ� ���Ȯ�� ����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'       - pAutoFg   : ������� ����
'       - pEqpCd    : ����ڵ�
'       - pVfyDt    : ���Ȯ������
'       - pVfyTm    : ���Ȯ�νð�
'       - pVfyId    : ���Ȯ����
'-----------------------------------------------------------------------------'
Public Function UpdateIISVfyDetail(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pAutoFg As String, _
                    ByVal pEqpCd As String, ByVal pVfyDt As String, ByVal pVfyTm As String, _
                    ByVal pVfyId As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS204 & " SET autofg='" & pAutoFg & "', eqpcd='" & pEqpCd & "'," & _
          "     vfydt='" & pVfyDt & "', vfytm='" & pVfyTm & "', vfyid='" & pVfyId & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateIISVfyDetail = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽����� ����ϴ� �����׸�, �����ڵ� ����Է� ����
'          ��������� Numeric�ΰ͸�
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'       - pRstVal   : �˻���1(Numeric)
'       - pRstCd    : �˻���2(Alphabetic)
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'       - pAutoFg   : ������� ����
'       - pRstType  : �������
'       - pEqpCd    : ����ڵ�
'-----------------------------------------------------------------------------'
Public Function UpdateIISResultNumeric(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstVal As Single, _
                    ByVal pRstCd As String, ByVal pHLDiv As String, ByVal pDPDiv As String, _
                    ByVal pAutoFg As String, ByVal pRstType As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    '## 1.1.8: �̻��(2005-04-11)
    '   - pRstVal�� �ڷ����� Long���� Single�� ����
    SQL = " UPDATE " & TIIS204 & " SET rstval=" & CStr(pRstVal) & ", rstcd='" & pRstCd & "'," & _
          "     hldiv='" & pHLDiv & "', dpdiv='" & pDPDiv & "', autofg='" & pAutoFg & "'," & _
          "     rsttype='" & pRstType & "', eqpcd='" & pEqpCd & "'," & _
          " transdt = TO_CHAR(sysdate, 'YYYYMMDD'), transtm = TO_CHAR(sysdate, 'HHMMSS') " & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateIISResultNumeric = SQL
End Function


'-----------------------------------------------------------------------------'
'   ��� : �������̽����� ����ϴ� �����׸�, �����ڵ� ����Է� ����
'          ��������� Free�ΰ͸�
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'       - pRstCd    : �˻���2(Alphabetic)
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'       - pAutoFg   : ������� ����
'       - pRstType  : �������
'       - pEqpCd    : ����ڵ�
'-----------------------------------------------------------------------------'
Public Function UpdateIISResultFree(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstCd As String, _
                    ByVal pHLDiv As String, ByVal pDPDiv As String, ByVal pAutoFg As String, _
                    ByVal pRstType As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS204 & " SET rstcd='" & pRstCd & "', hldiv='" & pHLDiv & "'," & _
          "     dpdiv='" & pDPDiv & "', autofg='" & pAutoFg & "'," & _
          "     rsttype='" & pRstType & "', eqpcd='" & pEqpCd & "'," & _
          " transdt = TO_CHAR(sysdate, 'YYYYMMDD'), transtm = TO_CHAR(sysdate, 'HHMMSS') " & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateIISResultFree = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽����� ����ϴ� �����׸�, �����ڵ� ���Ȯ�� ����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'       - pRstVal   : �˻���1(Numeric)
'       - pRstCd    : �˻���2(Alphabetic)
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'       - pAutoFg   : ������� ����
'       - pRstType  : �������
'       - pEqpCd    : ����ڵ�
'       - pVfyDt    : ���Ȯ������
'       - pVfyTm    : ���Ȯ�νð�
'       - pVfyId    : ���Ȯ����
'-----------------------------------------------------------------------------'
Public Function UpdateIISVfy(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstVal As Single, _
                    ByVal pRstCd As String, ByVal pHLDiv As String, ByVal pDPDiv As String, _
                    ByVal pAutoFg As String, ByVal pRstType As String, ByVal pEqpCd As String, _
                    ByVal pVfyDt As String, ByVal pVfyTm As String, ByVal pVfyId As String) As String
    Dim SQL As String
    
    '## 1.1.8: �̻��(2005-04-11)
    '   - pRstVal�� �ڷ����� Long���� Single�� ����
    SQL = " UPDATE " & TIIS204 & " SET rstval=" & CStr(pRstVal) & ", rstcd='" & pRstCd & "'," & _
          "     hldiv='" & pHLDiv & "', dpdiv='" & pDPDiv & "', autofg='" & pAutoFg & "'," & _
          "     rsttype='" & pRstType & "', eqpcd='" & pEqpCd & "', vfydt='" & pVfyDt & "'," & _
          "     vfytm='" & pVfyTm & "', vfyid='" & pVfyId & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateIISVfy = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� �Էµ� �˻��׸񰳼� ��������
'   �μ� :
'       - pSpcYy    : ��ü����
'       - pSpcNo    : ��ü��ȣ
'       - pInputCnt : �Էµ� �˻��׸� ����
'-----------------------------------------------------------------------------'
Public Function UpdateInputCnt(ByVal pSpcYy As String, ByVal pSpcNo As String, _
                    ByVal pInputCnt As Long) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS203 & " SET reqinputcnt=" & CStr(pInputCnt) & _
          " WHERE spcyy='" & pSpcYy & "' AND spcno=" & pSpcNo
          
    UpdateInputCnt = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� �Ѱ˻��װ���, �Էµ� �˻��׸񰳼� ��ȸ����
'   �μ� :
'       - pSpcYy    : ��ü����
'       - pSpcNo    : ��ü��ȣ
'-----------------------------------------------------------------------------'
Public Function SelectTotCnt(ByVal pSpcYy As String, ByVal pSpcNo As String)
    Dim SQL As String
    
    SQL = " SELECT reqtotcnt AS TOTCNT, reqinputcnt AS INPUTCNT" & _
          " FROM " & TIIS203 & " WHERE spcyy='" & pSpcYy & "' AND spcno=" & pSpcNo
          
    SelectTotCnt = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� ���Ȯ���ϴ� ����
'   �μ� :
'       - pSpcYy    : ��ü����
'       - pSpcNo    : ��ü��ȣ
'       - pStsCd    : ��ü����
'       - pVfyDt    : ���Ȯ������
'       - pVfyTm    : ���Ȯ�νð�
'       - pVfyId    : ���Ȯ����
'-----------------------------------------------------------------------------'
Public Function UpdateAccInfoVfy(ByVal pSpcYy As String, ByVal pSpcNo As String, _
                    ByVal pStsCd As String, ByVal pVfyDt As String, ByVal pVfyTm As String, _
                    ByVal pVfyId As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS203 & " SET stscd='" & pStsCd & "', vfydt='" & pVfyDt & "'," & _
          "     vfytm='" & pVfyTm & "', vfyid='" & pVfyId & "'" & _
          " WHERE spcyy='" & pSpcYy & "' AND spcno=" & pSpcNo
          
    UpdateAccInfoVfy = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ó�泻���� ���Ȯ���ϴ� ����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pStsCd    : ��ü����
'       - pExamDt   : ���Ȯ������
'       - pExamTm   : ���Ȯ�νð�
'       - pExamId   : ���Ȯ����
'-----------------------------------------------------------------------------'
Public Function UpdateOrderVfy(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pStsCd As String, ByVal pExamDt As String, _
                    ByVal pExamTm As String, ByVal pExamId As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS202 & " SET stscd='" & pStsCd & "', examdt='" & pExamDt & "'," & _
          "     examtm='" & pExamTm & "', examdoct='" & pExamId & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "' AND accseq=" & CStr(pAccSeq)
          
    UpdateOrderVfy = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb�� ������������ MAX(itemseq)+1�� ��ȸ����
'-----------------------------------------------------------------------------'
Public Function SelectMaxSeq() As String
    Dim SQL As String
    
    SQL = " SELECT MAX(itemseq)+1 AS MAXSEQ FROM " & TACC203
    
    SelectMaxSeq = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb�� �Ϲݰ˻� ���������� �Է��ϴ� ����
'   �μ� :
'       - pItemSeq : ���ۼ���
'       - pSpcNo   : ��ü��ȣ
'       - pPtId    : ȯ��ID
'       - pName    : �̸�
'       - pSex     : ����
'       - pAgeDay  : �Ϸ�
'       - pDoctNm  : ó���Ǹ�
'       - pDeptNm  : �������
'       - pWardNm  : ������
'       - pSpcCd   : ��ü�ڵ�
'       - pSpcNm   : ��ü��
'       - pStatFg  : ���޿���
'       - pQcFg    : QC����
'       - pSpcPos  : ��񿡼� ��ü��ġ
'       - pEqpCd   : ����ڵ�
'       - pTransDt : ��������
'       - pTransTm : ���۽ð�
'-----------------------------------------------------------------------------'
Public Function InsertClientAccInfo(ByVal pItemSeq As Long, ByVal pSpcNo As String, _
                    ByVal pPtId As String, ByVal pName As String, ByVal pSex As String, _
                    ByVal pAgeDay As Long, ByVal pDoctNm As String, ByVal pDeptNm As String, _
                    ByVal pWardNm As String, ByVal pSpcCd As String, ByVal pSpcNm As String, _
                    ByVal pStatFg As String, ByVal pQcFg As String, ByVal pSpcPos As String, _
                    ByVal pEqpCd As String, ByVal pTransDt As String, ByVal pTransTm As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TACC203 & " (itemseq, spcno, ptid, name, sex, ageday, doctnm, deptnm," & _
          "     wardnm, spccd, spcnm, statfg, qcfg, spcpos, eqpcd, transdt, transtm)" & _
          " VALUES (" & CStr(pItemSeq) & ", '" & pSpcNo & "', '" & pPtId & "', '" & pName & "'," & _
          "     '" & pSex & "', " & CStr(pAgeDay) & ", '" & pDoctNm & "', '" & pDeptNm & "'," & _
          "     '" & pWardNm & "', '" & pSpcCd & "', '" & pSpcNm & "', '" & pStatFg & "'," & _
          "     '" & pQcFg & "', '" & pSpcPos & "', '" & pEqpCd & "', '" & pTransDt & "', '" & pTransTm & "')"
    
    InsertClientAccInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb�� QC�˻� ���������� �Է��ϴ� ����
'   �μ� :
'       - pItemSeq : ���ۼ���
'       - pSpcNo   : ��ü��ȣ
'       - pPtId    : ȯ��ID
'       - pQcFg    : QC����
'       - pCtrlCd  : Control Code
'       - pLevelCd : Level Code
'       - pLotNo   : Lot No
'       - pSpcPos  : ��񿡼� ��ü��ġ
'       - pEqpCd   : ����ڵ�
'       - pTransDt : ��������
'       - pTransTm : ���۽ð�
'-----------------------------------------------------------------------------'
Public Function InsertClientQCAccInfo(ByVal pItemSeq As Long, ByVal pSpcNo As String, _
                    ByVal pPtId As String, ByVal pQcFg As String, ByVal pCtrlCd As String, _
                    ByVal pLevelCd As String, ByVal pLotNo As String, ByVal pSpcPos As String, _
                    ByVal pEqpCd As String, ByVal pTransDt As String, ByVal pTransTm As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TACC203 & " (itemseq, spcno, ptid, qcfg, ctrlcd, levelcd, lotno, spcpos," & _
          "     eqpcd, transdt, transtm)" & _
          " VALUES (" & CStr(pItemSeq) & ", '" & pSpcNo & "', '" & pPtId & "'," & _
          "     '" & pQcFg & "', '" & pCtrlCd & "', '" & pLevelCd & "', '" & pLotNo & "'," & _
          "     '" & pSpcPos & "', '" & pEqpCd & "', '" & pTransDt & "', '" & pTransTm & "')"
    
    InsertClientQCAccInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb�� ����� �Է��ϴ� ����
'   �μ� :
'       - pItemSeq  : ���ۼ���
'       - pIntNm    : ȭ��ǥ�� �˻��
'       - pTestCd   : �˻��ڵ�
'       - pResult   : ���
'       - pHLDiv    : H/L
'       - pDPDiv    : D/P
'-----------------------------------------------------------------------------'
Public Function InsertClientResult(ByVal pItemSeq As Long, ByVal pIntNm As String, _
                    ByVal pTestCd As String, ByVal pResult As String, ByVal pHLDiv As String, _
                    ByVal pDPDiv As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TACC204 & " (itemseq, intnm, testcd, result, hldiv, dpdiv)" & _
          " VALUES (" & CStr(pItemSeq) & ", '" & pIntNm & "', '" & pTestCd & "', '" & pResult & "'," & _
          "     '" & pHLDiv & "', '" & pDPDiv & "')"
          
    InsertClientResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb�� QC����� �Է��ϴ� ����
'   �μ� :
'       - pItemSeq  : ���ۼ���
'       - pIntNm    : ȭ��ǥ�� �˻��
'       - pTestCd   : �˻��ڵ�
'       - pResult   : ���
'       - pRADiv    : QC������� (A/R)
'       - pText     : QC Reject �Ұ�
'-----------------------------------------------------------------------------'
Public Function InsertClientQCResult(ByVal pItemSeq As Long, ByVal pIntNm As String, _
                    ByVal pTestCd As String, ByVal pResult As String, ByVal pRADiv As String, _
                    ByVal pText As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TACC204 & " (itemseq, intnm, testcd, result, radiv, comment)" & _
          " VALUES (" & CStr(pItemSeq) & ", '" & pIntNm & "', '" & pTestCd & "', '" & pResult & "'," & _
          "     '" & pRADiv & "', '" & pText & "')"
    InsertClientQCResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ���� ClientDb�� ������ ����
'   �μ� :
'       - pEqpCd  : ����ڵ�
'       - pBaseDt : ���� ������
'-----------------------------------------------------------------------------'
Public Function DeleteHisData(ByVal pEqpCd As String, ByVal pBaseDt As String) As String
    Dim SQL As String
    
    SQL = " DELETE FROM " & TACC203 & _
          " WHERE transdt<'" & pBaseDt & "' AND eqpcd='" & pEqpCd & "'"
    
    DeleteHisData = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��ü�� ��ȸ
'   �μ� :
'       - pSpcCd : ��ü�ڵ�
'-----------------------------------------------------------------------------'
Public Function SelectSpcNm(ByVal pSpcCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT FIELD3 AS SPCNM FROM " & TIIS002 & _
          " WHERE cdindex='" & CSPCCD & "' AND cdval1='" & pSpcCd & "'"
    
    SelectSpcNm = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������ȣ�� �̿��� ��ü����, ��ü������ ��ȸ
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'-----------------------------------------------------------------------------'
Public Function SelectBarNo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT spcyy AS SPCYY, spcno AS SPCNO FROM " & TIIS203 & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq)
          
    SelectBarNo = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش������ �˻��󸮽�Ʈ ��ȸ - 1.1.4: �̻��(2004-12-17)
'   �μ� :
'       - pFromDt : �������� From
'       - pToDt   : �������� To
'-----------------------------------------------------------------------------'
Public Function SelectTargetSpcs(ByVal pEqpCd As String, ByVal pFromDt As String, _
                                 ByVal pToDt As String) As String
    Dim SQL As String

    '## 1.2.2: �̻��(2005-06-14)
    '   - ó�渮��ũ�� ��ȸ�ϱ� ���� ��������
    '## �����׸�, �����ڵ��� ����Է��� �ȵȰ͸� ��ȸ!
    SQL = " SELECT a.spcyy AS SPCYY, a.spcno AS SPCNO, a.workarea AS WORKAREA, a.accdt AS ACCDT," & Chr(13) & _
          "        a.accseq AS ACCSEQ, a.ptid AS PTID, c." & FPTNM & " AS NAME, " & FJUMIN & " AS SSN," & Chr(13) & _
          "        a.sex AS SEX, a.statfg AS STATFG, a.wardid AS WARDID, a.deptcd AS DEPTCD," & Chr(13) & _
          "        d.field3 AS SPCNM, e.testnm AS TESTNM, f.empnm AS RCVNM, a.rcvdt AS RCVDT," & Chr(13) & _
          "        a.rcvtm AS RCVTM, g.mesg AS MESG" & Chr(13) & _
          " FROM " & TIIS203 & " a, " & TIIS204 & " b, " & THIS001 & " c, " & TIIS002 & " d," & Chr(13) & _
                TIIS101 & " e, " & THIS006 & " f, " & TIIS202 & " g" & Chr(13) & _
          " WHERE a.rcvdt BETWEEN '" & pFromDt & "' AND '" & pToDt & "' AND a.stscd IN ('2', '3')" & Chr(13) & _
          "     AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & Chr(13) & _
          "     AND b.testcd IN (SELECT TESTCD FROM " & TIIS404 & " WHERE eqpcd='" & pEqpCd & "')" & Chr(13) & _
          "     AND (b.rstcd IS NULL OR b.rstcd='')" & Chr(13) & _
          "     AND (b.vfydt IS NULL OR b.vfydt='')" & Chr(13) & _
          "     AND b.ptid=g.ptid AND b.orddt=g.orddt AND b.ordno=g.ordno AND b.ordseq=g.ordseq" & Chr(13) & _
          "     AND a.ptid=c." & FPTID & " AND d.cdindex='" & CSPCCD & "' AND a.spccd=d.cdval1" & Chr(13) & _
          "     AND b.testcd=e.testcd" & Chr(13) & _
          "     AND e.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=b.testcd)" & Chr(13) & _
          "     AND a.rcvid=f.empid(+)" & Chr(13) & _
          " ORDER BY a.workarea, a.accdt, a.accseq"
          
    SelectTargetSpcs = SQL
    
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش������ �˻��󸮽�Ʈ ��ȸ - 1.2.0: �̻��(2005-05-30)
'   �μ� :
'       - pEqpCd : ����ڵ�
'       - pDate  : ��ȸ��¥
'-----------------------------------------------------------------------------'
Public Function SelectTargetSpcsX(ByVal pEqpCd As String, ByVal pDate As String) As String
    Dim SQL As String
    
    '## NOTE: QUERY CHECK
    SQL = " SELECT DISTINCT a.spcyy AS SPCYY, a.spcno AS SPCNO," & _
          "     a.workarea+'-'+a.accdt+'-'+Convert(char(6), a.accseq) AS ACCNO," & _
          "     a.ptid AS PTID, c." & FPTNM & " AS NAME, a.rcvdt" & _
          " FROM " & TIIS203 & " a, " & TIIS204 & " b, " & THIS001 & " c" & _
          " WHERE a.rcvdt='" & pDate & "' AND a.workarea='" & IISCBCWA & "'" & _
          "     AND a.stscd IN ('2', '3', '4') AND (a.readfg='' OR a.readfg IS NULL)" & _
          "     AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
          "     AND b.testcd IN (SELECT testcd FROM " & TIIS404 & " WHERE eqpcd='" & pEqpCd & "')" & _
          "     AND a.ptid=c." & FPTID & _
          " ORDER BY a.rcvdt"
          
    SelectTargetSpcsX = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻��� ����Ʈ ��ȸ�� ��ȸ Flag ������Ʈ ���� - 1.2.0: �̻��(2005-05-30)
'-----------------------------------------------------------------------------'
Public Function UpdateTargetSpc(ByVal pSpcYy As String, ByVal pSpcNo As Long) As String
    Dim SQL As String
    
    '## NOTE: QUERY CHECK
    SQL = " UPDATE " & TIIS203 & " SET readfg='1'" & _
          " WHERE spcyy='" & pSpcYy & "' AND spcno=" & CStr(pSpcNo)
          
    UpdateTargetSpc = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽��� ����ϴ� ������� ��ȸ����(��ȿ����, D/P��ġ�� ��ȸ)
'          �������� ���� �׸� ��ȸ
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'-----------------------------------------------------------------------------'
Public Function SelectAutoRerunResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT a.testcd AS TESTCD, b.abbrnm10 AS TESTNM10, b.testnm AS TESTNM, c.field3 AS SPCNM," & _
          "     a.rstunit AS UNIT, a.lastrst AS LASTRST, b.rsttype AS RSTTYPE," & _
          "     a.rstdiv AS RSTDIV, a.detailfg AS DETAILFG, d.avalval AS AVALVAL," & _
          "     d.panicfg AS PANICFG, d.panicfrval AS PANICFRVAL, d.panictoval AS PANICTOVAL," & _
          "     d.deltafg AS DELTAFG, d.deltaval AS DELTAFRVAL, d.deltaval2 AS DELTATOVAL" & _
          " FROM " & TIIS204 & " a, " & TIIS101 & " b, " & TIIS002 & " c, " & TIIS102 & " d" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND (vfydt IS NULL OR vfydt='')" & _
          "     AND c.cdindex='" & CSPCCD & "' AND a.testcd=b.testcd" & _
          "     AND b.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
          "     AND a.spccd=c.cdval1" & _
          "     AND a.testcd=d.testcd AND a.spccd=d.spccd" & _
          "     AND d.applydt=(SELECT MAX(applydt) FROM " & TIIS102 & _
          "                    WHERE testcd=a.testcd AND spccd=a.spccd)" & _
          " ORDER BY b.rptseq"
          
    SelectAutoRerunResultInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽��� ����ϴ� Ư���˻� ������� ��ȸ���� - 1.0.1: �̻��(2005-10-11)
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'-----------------------------------------------------------------------------'
Public Function SelectSpecialResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT a.testcd AS TESTCD, b.abbrnm10 AS TESTNM10, b.testnm AS TESTNM" & _
          " FROM " & TIIS601 & " a, " & TIIS101 & " b" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND (a.vfydt IS NULL OR a.vfydt='')" & _
          "     AND a.testcd=b.testcd" & _
          "     AND b.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
          " ORDER BY b.rptseq"
    
    SelectSpecialResultInfo = SQL
End Function


'-----------------------------------------------------------------------------'
'   ��� : ������ȣ�� �̿��� ��ü����, ��ü������ ��ȸ
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'-----------------------------------------------------------------------------'
Public Function SelectBarNoByRack(ByVal pModuleNm As String, ByVal pRackNo As String, ByVal pRackPos As String) As String
    Dim SQL As String
    
    SQL = " SELECT SPC_NO " & _
          "   FROM PAM_RACKINFO " & _
          "  WHERE MODULE_NM = '" & pModuleNm & "'" & _
          "    AND RACK_NO = '" & pRackNo & "'" & _
          "    AND RACK_POS = '" & pRackPos & "'"
          
    SelectBarNoByRack = SQL

End Function

