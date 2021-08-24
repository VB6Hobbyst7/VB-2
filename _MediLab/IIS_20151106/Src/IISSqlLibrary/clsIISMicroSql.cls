VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMicroSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISMicroSql.cls (���ֿ�������)
'   �ۼ���  : �̻��
'   ��  ��  : �̻������� ����ϴ� ����
'   �ۼ���  : 2004-10-05
'   ��  ��  :
'       1. 1.0.2: �̻��(2004-12-04)
'       2. 1.0.3: �̻��(2004-12-06)
'       3. 1.1.6: �̻��(2005-01-03)
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   ��� : �������̽��� ����ϴ� �̻��� ������� ��ȸ����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCds  : ��� �˻��׸� ��ϵ� �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Function SelectMICResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCds As String) As String
    Dim SQL As String
    
    '## 1.0.2: �̻��(2004-12-04)
    '   - �Ѱ��� ������ȣ�� �������� Worksheet�� ���Եɼ� �־� Sensi�� ���� ��������� ��ȸ
    '     �ϵ��� ����
    '## 1.0.3: �̻��(2004-12-06)
    '   - ��� �˻��׸� ��ϵ� �˻��ڵ忡 �ش�Ǵ°͸� ��ȸ
    SQL = " SELECT a.wscd AS WSCD, a.wsunit AS WSUNIT, a.selfg AS SELFG, a.genfg AS GENFG," & _
          "     a.scfg AS SCFG, b.testcd AS TESTCD, c.abbrnm10 AS TESTNM10, c.testnm AS TESTNM," & _
          "     b.rsttype AS RSTTYPE, b.mfyseq AS MFYSEQ, b.ptid AS PTID" & _
          " FROM " & TIIS502 & " a, " & TIIS504 & " b, " & TIIS101 & " c" & _
          " WHERE a.workarea='" & pWorkarea & "' AND a.accdt='" & pAccDt & "'" & _
          "     AND a.accseq=" & CStr(pAccSeq) & " AND a.scfg IN ('" & IISGENSENSI & "','" & IISMIC & "')" & _
          "     AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
          "     AND b.stscd IN ('3','4') AND b.rsttype IN ('" & IISGENSENSI & "','" & IISMIC & "')" & _
          "     AND b.testcd IN (" & pTestCds & ")" & _
          "     AND b.testcd=c.testcd" & _
          "     AND c.applydt=(SELECT MAX(applydt) FROM " & TIIS101 & _
          "                    WHERE testcd=b.testcd AND applydt<=b.orddt)"
    
    SelectMICResultInfo = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : Vitek���� ������ �ո�(���)�� �̿� ���ڵ带 ��ȸ����
'   �μ� :
'       - pMnmCd : �ո�(���)
'-----------------------------------------------------------------------------'
Public Function SelectMnmCd(ByVal pMnmNm As String) As String
    Dim SQL As String

    SQL = " SELECT cdval1 AS MNMCD FROM " & TIIS002 & _
          " WHERE cdindex='" & CMNMCD & "' AND field1='" & pMnmNm & "'"
    
    SelectMnmCd = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �̻��� ������� ������Ʈ ����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pSenFg    : ��������� ����
'-----------------------------------------------------------------------------'
Public Function UpdateMicroResult(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pSenFg As String) As String
    Dim SQL As String
    
    '## 1.0.3: �̻��(2004-12-06)
    '   - ������ ����ڵ常 SenFg�� ������Ʈ
    SQL = " UPDATE " & TIIS504 & " SET senfg='" & pSenFg & "'" & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & _
          "     AND rsttype IN ('" & IISGENSENSI & "', '" & IISMIC & "')"
    
    UpdateMicroResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �̻��� �׻������ �Է� ����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCd   : �˻��ڵ�
'       - pMfySeq   : ����ȸ��
'       - pSeq      : Seq
'       - pMnmCd    : ���ڵ�
'       - pMicFg    : MIC����
'       - pMqtCd    : �����ڵ�
'       - pSCnt     : �׻��� �����
'       - pResult   : �׻��� ���
'-----------------------------------------------------------------------------'
Public Function InsertMICResult(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String, ByVal pMfySeq As Long, _
                    ByVal pSeq As Long, ByVal pMnmCd As String, ByVal pMicFg As String, _
                    ByVal pMqtCd As String, ByVal pSCnt As String, ByVal pResult As String) As String
    Dim strColumns As String    'INSERT���� �÷�����
    Dim SQL        As String
    Dim i          As Long
    
    '## �Է��� �׻������� ��ŭ �÷�����
    For i = 1 To pSCnt
        strColumns = strColumns & ", srst" & CStr(i)
    Next i
    
    SQL = " INSERT INTO " & TIIS505 & " (workarea, accdt, accseq, testcd, mfyseq, seq, mnmcd," & _
          "     micfg, mqtcd, scnt" & strColumns & ")" & _
          " VALUES ('" & pWorkarea & "', '" & pAccDt & "', " & CStr(pAccSeq) & "," & _
          "     '" & pTestCd & "', " & CStr(pMfySeq) & ", " & CStr(pSeq) & "," & _
          "     '" & pMnmCd & "', '" & pMicFg & "', '" & pMqtCd & "', " & CStr(pSCnt) & pResult & ")"
    
    InsertMICResult = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �׻�������� MAX(seq) ��ȸ ����
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'-----------------------------------------------------------------------------'
Public Function SelectMaxSeq(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT MAX(seq)+1 AS MAXSEQ FROM " & TIIS505 & _
          " WHERE workarea='" & pWorkarea & "' AND accdt='" & pAccDt & "'" & _
          "     AND accseq=" & CStr(pAccSeq)
    
    SelectMaxSeq = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : Vitek No�� �̿��� ��������(AccDt) ��ȸ���� - 1.1.6:�̻��(2005-01-03)
'   �μ� :
'       - pWorkarea : Workarea
'       - pMonth    : ������
'       - pAccSeq   : ��������
'-----------------------------------------------------------------------------'
Public Function SelectAccDt(ByVal pWorkarea As String, ByVal pMonth As String, _
                    ByVal pAccSeq As Long) As String
    Dim SQL As String
    
    SQL = " SELECT accdt FROM " & TIIS504 & _
          " WHERE workarea='" & pWorkarea & "' AND substr(accdt,5)='" & pMonth & "'" & _
          "     AND accseq=" & CStr(pAccSeq) & _
          " ORDER BY accdt DESC"
          
    SelectAccDt = SQL
End Function
