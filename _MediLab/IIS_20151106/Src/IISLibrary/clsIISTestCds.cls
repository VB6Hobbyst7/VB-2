VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTestCds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISTestCds.cls
'   �ۼ���  : �̻��
'   ��  ��  : �˻��ڵ� �÷��� Ŭ����
'   �ۼ���  : 2004-01-19
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mTestCds As Collection

Public Property Get Count() As Long
    Count = mTestCds.Count
End Property

Public Property Get Item(ByVal pKey As Variant) As clsIISTestCd
Attribute Item.VB_UserMemId = 0
    Set Item = mTestCds(pKey)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mTestCds.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ� �׸��߰�
'   �μ� :
'       1.pTestCd : �˻��ڵ�
'       2.pTestSeq: Sequence
'       3.pTestNm : �˻��
'   ��ȯ : �˻��ڵ� ��ü
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pTestCd As String, ByVal pTestSeq As Long, _
                    ByVal pTestNm As String) As clsIISTestCd
    Dim objTestCd As clsIISTestCd
    
    Set objTestCd = New clsIISTestCd
    With objTestCd
        .TestCd = pTestCd
        .TestSeq = pTestSeq
        .TestNm = pTestNm
    End With
    
    mTestCds.Add objTestCd, pTestCd
    
    Set Add = objTestCd
    Set objTestCd = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ� �׸��߰�(�÷��ǿ� Ű�� ���»��·� �߰�)
'   �μ� :
'       1.pTestCd : �÷��ǿ� �߰��� �˻��ڵ� Ŭ����
'       2.pBefore : �÷��ǿ��� ��ġ
'-----------------------------------------------------------------------------'
Public Sub AddEx(ByRef pTestCd As clsIISTestCd, Optional ByVal pBefore As Long = 0)
    If pBefore = 0 Then
        mTestCds.Add pTestCd
    Else
        mTestCds.Add pTestCd, , pBefore
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���޵� �÷��� Ŭ������ Ű�� �̿��� �Է�
'   �μ� :
'       1.pTestCds : Ű�� ���� �˻��ڵ� �÷��� Ŭ����
'-----------------------------------------------------------------------------'
Public Sub SortAdd(ByRef pTestCds As clsIISTestCds)
    Dim objTestCds As clsIISTestCds
    Dim objTestCd  As clsIISTestCd
    Dim i          As Long

    i = 1
    Set objTestCds = New clsIISTestCds
    For Each objTestCd In pTestCds
        objTestCds.Add objTestCd.TestCd, i, objTestCd.TestNm
        i = i + 1
    Next

    Set pTestCds = objTestCds
    Set objTestCds = Nothing
    Set objTestCd = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pKey As Variant)
    mTestCds.Remove pKey
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mTestCds.Count To 1 Step -1
        mTestCds.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    Dim objTestCd As clsIISTestCd
    
    For Each objTestCd In mTestCds
        If pTestCd = objTestCd.TestCd Then
            Exist = True
            GoTo EndLine
        End If
    Next
    Exist = False

EndLine:
    Set objTestCd = Nothing
End Function

Private Sub Class_Initialize()
    Set mTestCds = New Collection
End Sub

Private Sub Class_Terminate()
    Set mTestCds = Nothing
End Sub
