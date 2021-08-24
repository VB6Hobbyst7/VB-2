VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntResults.cls
'   �ۼ���  : �̻��
'   ��  ��  : �������̽� ��� �÷��� Ŭ����
'   �ۼ���  : 2004-05-26
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mIntResults As Collection

Public Property Get Count() As Long
    Count = mIntResults.Count
End Property

Public Property Get Item(ByVal pIntBase As String) As clsIISIntResult
Attribute Item.VB_UserMemId = 0
    Set Item = mIntResults(pIntBase)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mIntResults.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������̽� ����߰�1
'   �μ� :
'       - pIntBase   : ������ �˻��
'       - pIntNm     : ȭ��ǥ�� �˻��
'       - pIntResult : ��񿡼� ������ ���
'       - pResult    : ���
'       - pInfo      : �߰�����
'   ��ȯ : �������̽� �����ü
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pIntBase As String, ByVal pIntNm As String, ByVal pIntResult As String, _
                    ByVal pResult As String, Optional ByVal pInfo As String) As clsIISIntResult
    Dim objIntResult As clsIISIntResult
    
    Set objIntResult = New clsIISIntResult
    With objIntResult
        .IntBase = pIntBase
        .IntNm = pIntNm
        .IntResult = pIntResult
        .Result = pIntResult 'pResult
        .Info = pInfo
    End With
    mIntResults.Add objIntResult, pIntBase
    
    Set Add = objIntResult
    Set objIntResult = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽� ����߰�2
'   �μ� :
'       - pIntResult : �������̽� ���Ŭ����
'   ��ȯ : QC������� ��ü
'-----------------------------------------------------------------------------'
Public Sub AddX(ByVal pIntResult As clsIISIntResult)
    mIntResults.Add pIntResult, pIntResult.IntBase
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'   �μ� :
'       - pSeq : ���ۼ���
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pIntBase As String)
    mIntResults.Remove pIntBase
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mIntResults.Count To 1 Step -1
        mIntResults.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   �μ� :
'       - pSeq : ���ۼ���
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pIntBase As String) As Boolean
    Dim objIntResult As clsIISIntResult
    
    For Each objIntResult In mIntResults
        With objIntResult
            If pIntBase = .IntBase Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objIntResult = Nothing
End Function

Private Sub Class_Initialize()
    Set mIntResults = New Collection
End Sub

Private Sub Class_Terminate()
    Set mIntResults = Nothing
End Sub
