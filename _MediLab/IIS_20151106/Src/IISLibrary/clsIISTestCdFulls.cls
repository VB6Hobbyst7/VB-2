VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTestCdFulls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISTestCdFulls.cls
'   �ۼ���  : �̻��
'   ��  ��  : �˻��ڵ�(��ü) �÷��� Ŭ����
'   �ۼ���  : 2004-01-19
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mFTestCds As Collection      '�˻��ڵ� �÷���

Public Property Get Count() As Long
    Count = mFTestCds.Count
End Property

Public Property Get Item(ByVal pTestCd As String, ByVal pApplyDt As String) As clsIISTestCdFull
Attribute Item.VB_UserMemId = 0
    Set Item = mFTestCds(pTestCd & DIV & pApplyDt)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mFTestCds.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ� �׸��߰�
'   ��ȯ : �߰��� �˻��ڵ� ��ü
'   �μ� :
'       1.pTestCd    : �˻��ڵ�
'       2.pApplyDt   : ������
'       3.pExpireDt  : �����
'       4.pTestNm5   : �˻��(5�ڸ�)
'       5.pTestNm10  : �˻��(10�ڸ�)
'       6.pTestNm    : �˻��(��ü)
'       7.pWorkarea  : Workarea �ڵ�
'       8.pRstType   : �������(Null: �Ϲ�, F: Free, N: Numeric, A: Alpha)
'       9.pRstDiv    : Alternative, Require ����(A: Alternative, R: Require)
'       10.pPanelFg  : ó�汸��(Null: �����׸�, G: �׷��׸�, D: ���׸�)
'       11.pDetailFg : Detail �׸񿩺�(Null: ����, *:�����׸� ���ڵ�)
'       12.pRptSeq   : ��¼���
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pTestCd As String, ByVal pApplyDt As String, _
                    ByVal pWorkarea As String, ByVal pTestNm As String, _
                    Optional ByVal pExpireDt As String, Optional ByVal pTestNm5 As String, _
                    Optional ByVal pTestNm10 As String, Optional ByVal pRstType As String, _
                    Optional ByVal pRstDiv As String, Optional ByVal pPanelFg As String, _
                    Optional ByVal pDetailFg As String, Optional ByVal pRptSeq As Long _
                    ) As clsIISTestCdFull

    Dim objFTestCd As clsIISTestCdFull
    
    Set objFTestCd = New clsIISTestCdFull
    With objFTestCd
        .TestCd = pTestCd
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .TestNm5 = pTestNm5
        .TestNm10 = pTestNm10
        .TestNm = pTestNm
        .Workarea = pWorkarea
        .RstType = pRstType
        .RstDiv = pRstDiv
        .PanelFg = pPanelFg
        .DetailFg = pDetailFg
        .RptSeq = pRptSeq
    End With
    
    '## Ű�� �˻��ڵ�, ������
    mFTestCds.Add objFTestCd, pTestCd & DIV & pApplyDt
    
    Set Add = objFTestCd
    Set objFTestCd = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String, ByVal pApplyDt As String)
    mFTestCds.Remove pTestCd & DIV & pApplyDt
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mFTestCds.Count To 1 Step -1
        mFTestCds.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String, ByVal pApplyDt As String) As Boolean
    Dim objFTestCd As clsIISTestCdFull
    
    For Each objFTestCd In mFTestCds
        If pTestCd = objFTestCd.TestCd And pApplyDt = objFTestCd.Applydt Then
            Exist = True
            GoTo EndLine
        End If
    Next

EndLine:
    Set objFTestCd = Nothing
End Function

Public Function AddTestCd(ByRef pTestCdFull As clsIISTestCdFull)

End Function

Private Sub Class_Initialize()
    Set mFTestCds = New Collection
End Sub

Private Sub Class_Terminate()
    Set mFTestCds = Nothing
End Sub
