VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTSpcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISTSpcs.cls
'   �ۼ���  : �̻��
'   ��  ��  : ������ü �÷��� Ŭ����
'   �ۼ���  : 2004-01-20
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mTSpcs As Collection

Public Property Get Count() As Long
    Count = mTSpcs.Count
End Property

Public Property Get Item(ByVal pTestCd As String, ByVal pSpcCd As String, _
                         ByVal pApplyDt As String) As clsIISTSpc
Attribute Item.VB_UserMemId = 0
    Set Item = mTSpcs(pTestCd & DIV & pSpcCd & DIV & pApplyDt)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mTSpcs.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ� �׸��߰�
'   ��ȯ : ������ü ��ü
'   �μ� :
'       1.pTestCd       : �˻��ڵ�
'       2.pSpcCd        : ��ü�ڵ�
'       3.pSpcNm        : ��ü��
'       4.pApplyDt      : ������
'       5.pExpireDt     : �����
'       6.pUnit         : ����
'       7.pAvalVal      : ��ȿ����
'       8.pPanicFg      : Panic Check(0:No, 1:Yes)
'       9.pPanicFrVal   : Panic From Value
'       10.pPanicToVal  : Panic To Value
'       11.pDeltaFg     : Delta Check(0:No, 1:Yes)
'       12.pDeltaFrVal  : Delta From Value
'       13.pDeltaToVal  : Delta To Value
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pTestCd As String, ByVal pSpcCd As String, _
                    ByVal pSpcNm As String, ByVal pApplyDt As String, _
                    Optional pExpireDt As String, Optional pUnit As String, _
                    Optional pAvalVal As Long, Optional pPanicFg As String, _
                    Optional pPanicFrVal As Single, Optional pPanicToVal As Single, _
                    Optional pDeltaFg As String, Optional pDeltaFrVal As Long, _
                    Optional pDeltaToVal As Long) As clsIISTSpc

    Dim objTSpc As clsIISTSpc
    
    Set objTSpc = New clsIISTSpc
    With objTSpc
        .TestCd = pTestCd
        .SpcCd = pSpcCd
        .SpcNm = pSpcNm
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .Unit = pUnit
        .AvalVal = pAvalVal
        .PanicFg = pPanicFg
        .PanicFrVal = pPanicFrVal
        .PanicToVal = pPanicToVal
        .DeltaFg = pDeltaFg
        .DeltaFrVal = pDeltaFrVal
        .DeltaToVal = pDeltaToVal
    End With
    
    mTSpcs.Add objTSpc, pTestCd & DIV & pSpcCd & DIV & pApplyDt
    
    Set Add = objTSpc
    Set objTSpc = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String, ByVal pSpcCd As String, ByVal pApplyDt As String)
    mTSpcs.Remove pTestCd & DIV & pSpcCd & DIV & pApplyDt
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mTSpcs.Count To 1 Step -1
        mTSpcs.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String, ByVal pSpcCd As String, _
                      ByVal pApplyDt As String) As Boolean
    Dim objTSpc As clsIISTSpc
    
    For Each objTSpc In mTSpcs
        If pTestCd = objTSpc.TestCd And pSpcCd = objTSpc.SpcCd And pApplyDt = objTSpc.Applydt Then
            Exist = True
            GoTo EndLine
        End If
    Next
    Exist = False

EndLine:
    Set objTSpc = Nothing
End Function

Private Sub Class_Initialize()
    Set mTSpcs = New Collection
End Sub

Private Sub Class_Terminate()
    Set mTSpcs = Nothing
End Sub