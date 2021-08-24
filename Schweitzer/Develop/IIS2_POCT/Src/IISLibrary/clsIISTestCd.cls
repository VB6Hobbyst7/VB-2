VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTestCd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISTestCd.cls
'   �ۼ���  : �̻��
'   ��  ��  : �˻��ڵ� Ŭ����
'   �ۼ���  : 2004-01-19
'   ��  ��  :
'   ��  ��  : �˻��׸� �������� ����, ��������� ���� Ư���� ������� Ŭ����
'-----------------------------------------------------------------------------'

Option Explicit

Private mTestCd     As String       '�˻��ڵ�
Private mTestSeq    As Long         'Sequence
Private mTestNm     As String       '�˻��

Public Property Get Cd() As String
    Cd = mTestCd
End Property

Public Property Get CdNm() As String
    CdNm = mTestNm
End Property

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get TestSeq() As Long
    TestSeq = mTestSeq
End Property

Public Property Let TestSeq(ByVal vData As Long)
    mTestSeq = vData
End Property

Public Property Get TestNm() As String
    TestNm = mTestNm
End Property

Public Property Let TestNm(ByVal vData As String)
    mTestNm = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ���� ��ü�� ���޵� ��ü�� TestCd�� ��
'   �μ� :
'       1.pObj : clsIISTestCd ��ü
'   ��ȯ :
'       1.���簴ü < ���ް�ü �̸� -1
'       2.���簳ü = ���ް�ü �̸� 0
'       3.���簳ü > ���ް�ü �̸� 1
'-----------------------------------------------------------------------------'
Public Function Compare(ByVal pObj As Object)
    Compare = StrComp(mTestCd, pObj.TestCd, vbTextCompare)
End Function