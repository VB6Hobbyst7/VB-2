VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISQcSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISQcSql.cls (���ֿ�������)
'   �ۼ���  : �̻��
'   ��  ��  : QC���� ����ϴ� ����
'   �ۼ���  : 2004-10-05
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   ��� : �������̽��� ����ϴ� QC������� ��ȸ����(��ȿ����, D/P��ġ�� ��ȸ)
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'-----------------------------------------------------------------------------'
Public Function SelectQCResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    '## ������ ����� ����
'    SQL = " SELECT a.testcd AS TESTCD, c.abbrnm10 AS TESTNM10, c.testnm AS TESTNM, a.rstunit AS UNIT," & _
'          "     a.rstdiv AS RSTDIV, a.ctrlcd AS CTRLCD, a.levelcd AS LEVELCD, a.lotno AS LOTNO," & _
'          "     b.meanval AS MEANVAL, b.sdval AS SDVAL, b.avalval AS AVALVAL, b.refcd AS REFCD," & _
'          "     b.cvval AS CVVAL, b.minval AS MINVAL, b.maxval AS MAXVAL, b.wmset AS WMSET" & _
'          " FROM " & TIIS307 & " a, " & TIIS304 & " b, " & TIIS101 & " c" & _
'          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
'          "     AND a.accseq=" & CStr(pAccSeq) & " AND (a.vfydt='' OR a.vfydt IS NULL)" & _
'          "     AND a.ctrlcd=b.ctrlcd AND a.levelcd=b.levelcd AND a.lotno=b.lotno AND a.testcd=b.testcd" & _
'          "     AND a.testcd=c.testcd" & _
'          "     AND c.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
'          " ORDER BY c.rptseq"
          
    '## ���� ����� ����
    SQL = " SELECT a.testcd AS TESTCD, c.abbrnm10 AS TESTNM10, c.testnm AS TESTNM, a.rstunit AS UNIT," & _
          "     a.rstdiv AS RSTDIV, a.ctrlcd AS CTRLCD, a.levelcd AS LEVELCD, a.lotno AS LOTNO," & _
          "     b.meanval AS MEANVAL, b.sdval AS SDVAL, b.avalval AS AVALVAL, b.refcd AS REFCD," & _
          "     b.cvval AS CVVAL, b.minval AS MINVAL, b.maxval AS MAXVAL, b.wmset AS WMSET" & _
          " FROM " & TIIS307 & " a, " & TIIS304 & " b, " & TIIS101 & " c" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND (a.vfydt='' OR a.vfydt IS NULL)" & _
          "     AND (a.rstcd='' OR a.rstcd IS NULL)" & _
          "     AND a.ctrlcd=b.ctrlcd AND a.levelcd=b.levelcd AND a.lotno=b.lotno AND a.testcd=b.testcd" & _
          "     AND a.testcd=c.testcd" & _
          "     AND c.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & " WHERE testcd=a.testcd)" & _
          " ORDER BY c.rptseq"
    
    SelectQCResultInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش���Ʈ��, �˻��ڵ��� ���Ű�� 9�� ��ȸ����
'   �μ� :
'       - pCtrlCd  : Control Code
'       - pLevelCd : Level Code
'       - pLotNo   : Lot No
'       - pTestCd  : �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Function SelectHistoryResult(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                    ByVal pLotNo As String, ByVal pTestCd As String) As String
    Dim SQL As String
    
    SQL = " SELECT rstcd AS RSTCD, rsttype AS RSTTYPE, radiv AS RADIV" & _
          " FROM (SELECT * FROM " & TIIS307 & _
          "       WHERE ctrlcd='" & pCtrlCd & "' AND levelcd='" & pLevelCd & "' AND lotno='" & pLotNo & "'" & _
          "             AND testcd='" & pTestCd & "' AND rsttype='N' AND (vfydt<>'' OR vfydt IS NOT NULL)" & _
          "       ORDER BY vfydt DESC, vfytm DESC)" & _
          " WHERE rownum<=9"
    
    SelectHistoryResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽����� QC ����Է� ���� (��������� Numeric�ΰ͸�)
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'       - pRstVal   : �˻���1(Numeric)
'       - pRstCd    : �˻���2(Alphabetic)
'       - pRstType  : ������� (F:Free, N:Numeric, A:Alpha)
'       - pRADiv    : ������� (A:Accept, R:Reject)
'       - pTxtFg    : Text��� ����(Null:��, 1:��)
'       - pAutoFg   : ������� ����
'       - pEqpCd    : ����ڵ�
'-----------------------------------------------------------------------------'
Public Function UpdateQCResultNumeric(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstVal As Single, _
                    ByVal pRstCd As String, ByVal pRstType As String, ByVal pRADiv As String, _
                    ByVal pTxtFg As String, ByVal pAutoFg As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS307 & " SET rstval=" & CStr(pRstVal) & ", rstcd='" & pRstCd & "'," & _
          "     rsttype='" & pRstType & "', radiv='" & pRADiv & "', txtfg='" & pTxtFg & "'," & _
          "     autofg='" & pAutoFg & "', eqpcd='" & pEqpCd & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"

    UpdateQCResultNumeric = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽����� QC ����Է� ���� (��������� Free�ΰ͸�)
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'       - pRstCd    : �˻���2(Alphabetic)
'       - pRstType  : ������� (F:Free, N:Numeric, A:Alpha)
'       - pRADiv    : ������� (A:Accept, R:Reject)
'       - pTxtFg    : Text��� ����(Null:��, 1:��)
'       - pAutoFg   : ������� ����
'       - pEqpCd    : ����ڵ�
'-----------------------------------------------------------------------------'
Public Function UpdateQCResultFree(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pRstCd As String, _
                    ByVal pRstType As String, ByVal pRADiv As String, ByVal pTxtFg As String, _
                    ByVal pAutoFg As String, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS307 & " SET rstcd='" & pRstCd & "', rsttype='" & pRstType & "'," & _
          "     radiv='" & pRADiv & "', txtfg='" & pTxtFg & "', autofg='" & pAutoFg & "'," & _
          "     eqpcd='" & pEqpCd & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    UpdateQCResultFree = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽����� QC �Ұ߰�� ��ȸ����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Function SelectQCComment(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String) As String
    Dim SQL As String
    
    '## DOIT: QUERY CHECK
    SQL = " SELECT seq AS SEQ FROM " & TIIS308 & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "'"
    
    SelectQCComment = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽����� QC �Ұ߰�� �Է�����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Function InsertQCComment(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pSeq As Long, _
                    ByVal pText As String) As String
    Dim SQL As String
    
    SQL = " INSERT INTO " & TIIS308 & " (workarea, accdt, accseq, testcd, seq, text)" & _
          "     VALUES ('" & pWorkarea & "', '" & pAccDt & "', " & CStr(pAccSeq) & "," & _
          "         '" & pTestCd & "', " & CStr(pSeq) & ", '" & pText & "')"
    
    InsertQCComment = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽����� QC �Ұ߰�� ��������
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Function UpdateQCComment(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pSeq As Long, _
                    ByVal pText As String) As String
    Dim SQL As String
    
    SQL = " UPDATE " & TIIS308 & " SET text='" & pText & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & " AND testcd='" & pTestCd & "' AND seq=" & CStr(pSeq)
    
    UpdateQCComment = SQL
End Function