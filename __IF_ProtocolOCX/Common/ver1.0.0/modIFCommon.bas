Attribute VB_Name = "modIFCommon"
Option Explicit

'Sample �� ���� ����
Type SAMPLE_INFO
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    ORDCNT  As Integer
    IFCD()  As String
    SVOL()  As String
    KIND    As String
    SINDEX  As Boolean
End Type

'���ŵ� ����� ���� ����
Type RESULT_INFO
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    RSTCNT  As Integer
    IFCD    As String
    RST1    As String
    RST2    As String
    UNIT    As String
    FLAG    As String
    ALARMCD As String
    RSTGBN  As String
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
