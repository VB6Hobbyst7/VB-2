VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISBuffers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISBuffers.cls
'   �ۼ���  : �̻��
'   ��  ��  : ���� �÷��� Ŭ����
'   �ۼ���  : 2004-04-28
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBuffers As Collection

Public Property Get Count() As Long
    Count = mBuffers.Count
End Property

Public Property Get Item(ByVal pSeq As Long) As clsIISBuffer
Attribute Item.VB_UserMemId = 0
    Set Item = mBuffers(CStr(pSeq))
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mBuffers.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : ���۳��� �߰�
'   �μ� :
'       - pSeq      : ���ۼ���
'       - pBuffers  : ���۳���
'   ��ȯ : ���� Ŭ����
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pSeq As Long, ByVal pBuffers As String) As clsIISBuffer
    Dim objBuffer As clsIISBuffer
    
    Set objBuffer = New clsIISBuffer
    With objBuffer
        .Seq = pSeq
        .Buffers = pBuffers
    End With
    
    mBuffers.Add objBuffer, CStr(pSeq)
    
    Set Add = objBuffer
    Set objBuffer = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'   �μ� :
'       - pSeq : ���ۼ���
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pSeq As Long)
    mBuffers.Remove CStr(pSeq)
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mBuffers.Count To 1 Step -1
        mBuffers.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   �μ� :
'       - pSeq : ���ۼ���
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pSeq As Long) As Boolean
    Dim objBuffer As clsIISBuffer
    
    For Each objBuffer In mBuffers
        With objBuffer
            If pSeq = .Seq Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objBuffer = Nothing
End Function

Private Sub Class_Initialize()
    Set mBuffers = New Collection
End Sub

Private Sub Class_Terminate()
    Set mBuffers = Nothing
End Sub
