Attribute VB_Name = "modIFCommon"
Option Explicit

'Sample �� ���� ����
Type SAMPLE_INFO
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    KIND    As String       '1st/Rerun ����
    ORDCNT  As Integer
    IFCD()  As String
    SVOL()  As String
    SINDEX  As Boolean      'Serum Index �˻翩��
    SPCCD   As String       '��ü����(Rack Type)
    RSTDT   As String       '2005/6/10 yk
    OTHER   As String
    CMT1    As String       '2005/8/1 yk
    INSTID  As String       '2006/2/9 yk
    INSTNM  As String       '2006/10/11 yk
    CONTROLNAME As String   '2007/01/17    by YeounJu
    CONTAINER   As String   'Container Type(for Cobas6000)...2007/6/22 yk
    '<2008/03/10 mc
    PatName As String
    Sex     As String
    Age     As String
    HosName As String
    Total   As String
    '>
End Type

'���ŵ� ����� ���� ����
Type RESULT_INFO
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    KIND    As String
    RSTCNT  As Integer
    IFCD    As String
    RST1    As String
    RST2    As String
    UNIT    As String
    FLAG    As String
    ALARMCD As String
    INSTID  As String
    INSTNM  As String       '2006/10/11 yk
    RSTDT   As String       'Date/Time Results Reported or Last Modified (2005/6/10 �߰� yk)
    SPCCD   As String       '��ü���� (2006/10/11 �߰� yk)
    OPERID  As String       'Operation ID (       "      )
    OTHER   As String
End Type


'�ӽ÷� ��й�ȣ ����
Public Const pOpenPW = "ACK"
Public Const pEditPW = "MEDI@CK"
'
'   ASTM Protocol CheckSum ���
'
Public Function ChkSum_ASTM(ByVal Para As String) As String

    Dim i   As Integer
    Dim Tmp As Integer
    Dim ChkS1   As Integer
    Dim ChkS2   As String
    
    For i = 1 To Len(Para)
        Tmp = Asc(Mid$(Para, i, 1))
        ChkS1 = ChkS1 + Tmp
    Next i
    ChkS1 = ChkS1 Mod 256
    ChkS2 = Right$("0" & Hex$(ChkS1), 2)
    
    ChkSum_ASTM = ChkS2
    
End Function



