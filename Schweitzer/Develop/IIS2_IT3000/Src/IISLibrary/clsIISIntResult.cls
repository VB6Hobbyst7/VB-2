VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntResult.cls
'   �ۼ���  : �̻��
'   ��  ��  : �������̽� ��� Ŭ����
'   �ۼ���  : 2004-05-26
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mIntBase    As String     '������ �˻��
Private mIntNm      As String     'ȭ��ǥ�� �˻��
Private mIntResult  As String     '��񿡼� ������ ���
Private mResult     As String     'IIS���� ������ ���
Private mInfo       As String     '��񿡼� ������ �߰�����

Public Property Get IntBase() As String
    IntBase = mIntBase
End Property

Public Property Let IntBase(ByVal vData As String)
    mIntBase = vData
End Property

Public Property Get IntNm() As String
    IntNm = mIntNm
End Property

Public Property Let IntNm(ByVal vData As String)
    mIntNm = vData
End Property

Public Property Get IntResult() As String
    IntResult = mIntResult
End Property

Public Property Let IntResult(ByVal vData As String)
    mIntResult = vData
End Property

Public Property Get Result() As String
    Result = mResult
End Property

Public Property Let Result(ByVal vData As String)
    mResult = vData
End Property

Public Property Get Info() As String
    Info = mInfo
End Property

Public Property Let Info(ByVal vData As String)
    mInfo = vData
End Property