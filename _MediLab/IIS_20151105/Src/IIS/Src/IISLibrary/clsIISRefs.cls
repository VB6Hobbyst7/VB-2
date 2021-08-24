VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISRefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISRefs.cls
'   �ۼ���  : �̻��
'   ��  ��  : ����ġ �÷��� Ŭ����
'   �ۼ���  : 2004-01-20
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mRefs As Collection

Public Property Get Count() As Long
    Count = mRefs.Count
End Property

Public Property Get Item(ByVal pTestCd As String, ByVal pSpcCd As String, _
                         ByVal pSex As String, ByVal pAgeFr As Long, _
                         ByVal pAgeTo As Long, ByVal pApplyDt As String) As clsIISRef
Attribute Item.VB_UserMemId = 0
                         
    Set Item = mRefs(pTestCd & DIV & pSpcCd & DIV & pSex & DIV & CStr(pAgeFr) & DIV & _
                    CStr(pAgeTo) & DIV & pApplyDt)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mRefs.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ� �׸��߰�
'   ��ȯ : ������ü ��ü
'   �μ� :
'       1.pTestCd   : �˻��ڵ�
'       2.pSpcCd    : ��ü�ڵ�
'       3.pSex      : ����
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : ������
'       7.pExpireDt : �����
'       8.pRefFrVal : From Reference Value
'       9.pRefToVal : To Reference Value
'       10.pRefCd   : Alpha��� ����ġ
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pTestCd As String, ByVal pSpcCd As String, _
                    ByVal pSex As String, ByVal pAgeFr As Long, _
                    ByVal pAgeTo As Long, ByVal pApplyDt As String, _
                    Optional ByVal pExpireDt As String, Optional ByVal pRefFrVal As Single, _
                    Optional ByVal pRefToVal As Single, Optional ByVal pRefCd As String) As clsIISRef

    Dim objTSpc As clsIISRef
    
    Set objTSpc = New clsIISRef
    With objTSpc
        .TestCd = pTestCd
        .SpcCd = pSpcCd
        .Sex = pSex
        .AgeFr = pAgeFr
        .AgeTo = pAgeTo
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .RefFrVal = pRefFrVal
        .RefToVal = pRefToVal
        .Refcd = pRefCd
    End With
    
    mRefs.Add objTSpc, pTestCd & DIV & pSpcCd & DIV & pSex & DIV & CStr(pAgeFr) & DIV & _
                CStr(pAgeTo) & DIV & pApplyDt
    
    Set Add = objTSpc
    Set objTSpc = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String, ByVal pSpcCd As String, _
                  ByVal pSex As String, ByVal pAgeFr As Long, _
                  ByVal pAgeTo As Long, ByVal pApplyDt As String)
                         
    mRefs.Remove pTestCd & DIV & pSpcCd & DIV & pSex & DIV & CStr(pAgeFr) & DIV & _
                    CStr(pAgeTo) & DIV & pApplyDt
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mRefs.Count To 1 Step -1
        mRefs.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String, ByVal pSpcCd As String, _
                         ByVal pSex As String, ByVal pAgeFr As Long, _
                         ByVal pAgeTo As Long, ByVal pApplyDt As String) As Boolean
    Dim objTSpc As clsIISRef
    
    For Each objTSpc In mRefs
        With objTSpc
            If pTestCd = .TestCd And pSpcCd = .SpcCd And pSex = .Sex And pAgeFr = .AgeFr And _
               pAgeTo = .AgeTo And pApplyDt = .Applydt Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objTSpc = Nothing
End Function

Private Sub Class_Initialize()
    Set mRefs = New Collection
End Sub

Private Sub Class_Terminate()
    Set mRefs = Nothing
End Sub
