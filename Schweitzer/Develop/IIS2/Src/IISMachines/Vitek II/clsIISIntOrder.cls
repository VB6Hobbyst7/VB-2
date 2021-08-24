VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntOrder.cls
'   �ۼ���  : �̻��
'   ��  ��  : Vitek ��� �������� Ŭ����
'   �ۼ���  : 2005-01-31
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mCheckSum    As String       'Check Sum
Private mSeq         As Long         '������ ������ Spread�� Row
Private mSendCnt     As Long         '������ ��ü����

Public Property Get CheckSum() As String
    CheckSum = mCheckSum
End Property

Public Property Let CheckSum(ByVal vData As String)
    mCheckSum = vData
End Property

Public Property Get Seq() As Long
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As Long)
    mSeq = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    mSendCnt = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������ڿ� ��ȸ
'   �μ� :
'       - pAccInfo  : �������� Ŭ����
'       - pVitekNo  : Vitek No
'   ��ȯ : �������ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo, ByVal pVitekNo As String) As String
    Dim strSexAge As String  '�۽��� Sex/Age
    Dim strColTm  As String  '�۽��� ä���ð� (HH:MM����)
    Dim strRcvTm  As String  '�۽��� �����ð� (HH:MM����)
    Dim strOrder  As String  '�������� ���ڿ�
    Dim strColDt  As String
    
    '## px(�����), ss(��ü), ct(�˻�����) �ʵ�� Vitek�� �̸� ���ǵǾ�� ��!!
    '   ct�ʵ�� �׻� cs�� Culture & Sensitivity�� ����ϰ� �ؾ���!
    With pAccInfo
        '-- 2010.07.05 osw : ����/���̰� ���� ���� �߻� ==> �ȵ�����
        'strSexAge = .Sex & "/" & Mid$(.Age, 1, Len(.Age) - 1)
        strSexAge = ""
        strColTm = Format$(Format$(.ColTm, "##:##:##"), "HH:MM")
        strRcvTm = Format$(Format$(.ColTm, "##:##:##"), "HH:MM")
        
        '## ��� ����Ұ�!
        'strOrder = "mtmpr|pi" & .PtId & "|w1" & strSexAge & "|px" & .DeptCd & "|pl" & .WardId & _
                   "|si|ss" & .SpcCd & "|s1" & .ColDt & "|s2" & strColTm & "|s3" & .RcvDt & _
                   "|s4" & strRcvTm & "|ci" & pVitekNo & "|ctcs|zz|" & vbCrLf
    
        '## ��¥���˺��� : �ȵ�����
        strColDt = Mid(.ColDt, 1, 4) & "/" & Mid(.ColDt, 5, 2) & "/" & Mid(.ColDt, 7, 2)
        
        strOrder = "mtmpr|pi" & .PtId & "|w1" & strSexAge & "|px" & .DeptCd & "|pl" & .WardId & _
                   "|si|ss" & .SpcCd & "|s1" & strColDt & "|s2" & strColTm & "|s3" & .RcvDt & _
                   "|s4" & strRcvTm & "|ci" & pVitekNo & "|ctcs|zz|" & vbCrLf
    
    End With
    
    mCheckSum = GetChkSum(strOrder & "")
    GetOrder = strOrder
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڿ��� CheckSum�� ����
'   �μ� :
'       - pMsg : ���ڿ�
'   ��ȯ : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = LCase(Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2))
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mCheckSum = ""
    mSeq = 0
End Sub