VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntNms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntNms.cls
'   �ۼ���  : �̻��
'   ��  ��  : ��� �˻��׸� �÷��� Ŭ����
'   �ۼ���  : 2004-04-29
'   ��  ��  :
'       1. 1.1.13: �̻��(2005-09-20)
'-----------------------------------------------------------------------------'

Option Explicit

Private mIntNms As Collection

Public Property Get Count() As Long
    Count = mIntNms.Count
End Property

Public Property Get Item(ByVal pTestCd As String) As clsIISIntNm
Attribute Item.VB_UserMemId = 0
    Set Item = mIntNms(pTestCd)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mIntNms.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : �˻��ڵ� �߰� (�������̽���)
'          �˻��ڵ��� ������ �˻���� ��ȸ�ϱ� ����
'   �μ� :
'       - pTestCd   : �˻��ڵ�
'       - pIntNm    : ȭ��ǥ�� �˻��
'       - pIntBase  : ������ �˻��
'       - pFrVal    : BorderLine From Value
'       - ptoVal    : BorderLine To Value
'   ��ȯ : �˻��׸� Ŭ����
'-----------------------------------------------------------------------------'
Public Function AddIIS(ByVal pTestCd As String, ByVal pIntNm As String, ByVal pIntBase As String, _
                    Optional ByVal pFrVal As String, Optional ByVal pToVal As String) As clsIISIntNm
    Dim objIntNm As clsIISIntNm
    
    Set objIntNm = New clsIISIntNm
    With objIntNm
        .TestCd = pTestCd
        .IntNm = pIntNm
        .IntBase = pIntBase
        .FrVal = pFrVal
        .ToVal = pToVal
    End With
    
    mIntNms.Add objIntNm, pTestCd
    
    Set AddIIS = objIntNm
    Set objIntNm = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'   �μ� :
'       - pTestCd   : �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String)
    mIntNms.Remove pTestCd
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mIntNms.Count To 1 Step -1
        mIntNms.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   �μ� :
'       - pTestCd   : �˻��ڵ�
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    Dim objIntNm As clsIISIntNm     '��� �˻��׸� Ŭ����
    
    For Each objIntNm In mIntNms
        With objIntNm
            If pTestCd = .TestCd Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next

EndLine:
    Set objIntNm = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������ �˻���� �̿��� �˻�� ��ȸ
'   �μ� :
'       - pIntBase : ������ �˻��
'   ��ȯ : �˻��
'-----------------------------------------------------------------------------'
Public Function GetIntNm(ByVal pIntBase As String) As String
    Dim objIntNm As clsIISIntNm     '��� �˻��׸� Ŭ����
    
    For Each objIntNm In mIntNms
        If objIntNm.IntBase = pIntBase Then
            GetIntNm = objIntNm.IntNm
            GoTo EndLine
        End If
    Next
    
EndLine:
    Set objIntNm = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻���� �̿��� ������ �˻�� ��ȸ - 1.1.13: �̻��(2005-09-20)
'   �μ� :
'       - pIntNm : �˻��
'   ��ȯ : �˻��
'-----------------------------------------------------------------------------'
Public Function GetIntBase(ByVal pIntNm As String) As String
    Dim objIntNm As clsIISIntNm     '��� �˻��׸� Ŭ����
    
    For Each objIntNm In mIntNms
        If objIntNm.IntNm = pIntNm Then
            GetIntBase = objIntNm.IntBase
            GoTo EndLine
        End If
    Next
    
EndLine:
    Set objIntNm = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������ �˻���� �÷��ǿ� �ִ��� ��ȸ
'   �μ� :
'       - pIntBase : ������ �˻��
'   ��ȯ : True(����), False(������)
'-----------------------------------------------------------------------------'
Public Function ExistIntBase(ByVal pIntBase As String) As Boolean
    Dim objIntNm As clsIISIntNm     '��� �˻��׸� Ŭ����
    
    For Each objIntNm In mIntNms
        If objIntNm.IntBase = pIntBase Then
            ExistIntBase = True
            GoTo EndLine
        End If
    Next
    
EndLine:
    Set objIntNm = Nothing
End Function

Private Sub Class_Initialize()
    Set mIntNms = New Collection
End Sub

Private Sub Class_Terminate()
    Set mIntNms = Nothing
End Sub