VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISResults.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : ������� �÷��� Ŭ����
'   �ۼ���  : 2004-03-31
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mResults As Collection

Public Property Get Count() As Long
    Count = mResults.Count
End Property

Public Property Get Item(ByVal pTestCd As String) As clsIISResult
Attribute Item.VB_UserMemId = 0
    Set Item = mResults(pTestCd)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mResults.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : ������� �߰�1 (�������̽���)
'   �μ� :
'       - pTestCd   : �˻��ڵ�
'       - pTestNm10 : �˻��(10�ڸ�)
'       - pTestNm   : �˻�� (��ü)
'       - pUnit     : �������
'       - pRerunFg  : ��˿��� (0:��, 1:��)
'       - pRstType  : ������� (F:Free, N:Numeric, A:Alpha)
'       - pRstDiv   : Require/Alternative ���� (R:Require, A:Alternative, *:�󼼸��ڵ�)
'       - pDetailFg : ���׸����
'       - pTSpc     : ������ü ��ü
'   ��ȯ : ������� ��ü
'-----------------------------------------------------------------------------'
Public Function AddInter(ByVal pTestCd As String, Optional ByVal pTestNm10 As String, _
                    Optional ByVal pTestNm As String, Optional ByVal pUnit As String, _
                    Optional ByVal pLastRst As String, Optional ByVal pRerunFg As String, _
                    Optional ByVal pRstType As String, Optional ByVal pRstDiv As String, _
                    Optional ByVal pDetailFg As String, Optional ByVal pTSpc As clsIISTSpc) As clsIISResult
    Dim objResult As clsIISResult
    
    Set objResult = New clsIISResult
    With objResult
        .TestCd = pTestCd
        .TestNm10 = pTestNm10
        .TestNm = pTestNm
        .Unit = pUnit
        .LastRst = pLastRst
        .RerunFg = pRerunFg
        .RstType = pRstType
        .RstDiv = pRstDiv
        .DetailFg = pDetailFg
        Set .TSpc = pTSpc
    End With
    
    mResults.Add objResult, pTestCd
    
    Set AddInter = objResult
    Set objResult = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������� �߰�2 (�������̽���)
'   �μ� :
'       - pResult : ������� ��ü
'   ��ȯ : ������� ��ü
'-----------------------------------------------------------------------------'
Public Sub AddInterX(ByVal pResult As clsIISResult)
    mResults.Add pResult, pResult.TestCd
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'   �μ� :
'       - pTestCd : �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String)
    mResults.Remove pTestCd
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mResults.Count To 1 Step -1
        mResults.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   �μ� :
'       - pTestCd : �˻��ڵ�
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    Dim objResult As clsIISResult
    
    For Each objResult In mResults
        With objResult
            If pTestCd = .TestCd Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objResult = Nothing
End Function

Private Sub Class_Initialize()
    Set mResults = New Collection
End Sub

Private Sub Class_Terminate()
    Set mResults = Nothing
End Sub
