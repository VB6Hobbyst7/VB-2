VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntErrors.cls
'   �ۼ���  : �̻��
'   ��  ��  : �������̽��� �߻��ϴ� �������� �÷��� Ŭ����
'   �ۼ���  : 2004-05-09
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mIntErrs As Collection

Public Property Get Count() As Long
    Count = mIntErrs.Count
End Property

Public Property Get Item(ByVal pSeq As Long) As clsIISIntError
Attribute Item.VB_UserMemId = 0
    Set Item = mIntErrs(CStr(pSeq))
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mIntErrs.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� �߰�(�����, ��ü���ÿ���)
'   �μ� :
'       - pCode  : �����ڵ�
'       - pEqpCd : ����ڵ�
'       - pEqpNm : ����
'       - pBarNo : ���ڵ��ȣ
'       - pPtId  : ȯ��ID
'       - pName  : ȯ�ڸ�
'   ��ȯ : ������ü
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pCode As String, Optional pEqpCd As String, Optional pEqpNm As String, _
                    Optional ByVal pBarNo As String, Optional ByVal pPtId As String, _
                    Optional ByVal pName As String) As clsIISIntError
    Dim objIntErr As clsIISIntError
    Dim lngSeq    As Long       '��������
    
    lngSeq = mIntErrs.Count + 1
    Set objIntErr = New clsIISIntError
    With objIntErr
        .Seq = lngSeq
        .Code = pCode
        .ErrDt = Format$(Now, "YYYY-MM-DD HH:MM:SS")
        .EqpCd = pEqpCd
        .EqpNm = pEqpNm
        .BarNo = pBarNo
        .PtId = pPtId
        .Name = pName
    End With
    
    mIntErrs.Add objIntErr, CStr(lngSeq)
    
    Set Add = objIntErr
    Set objIntErr = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� �߰�(��񿡼� ������ ����)
'   �μ� :
'       - pCode  : �����ڵ�
'       - pEqpCd : ����ڵ�
'       - pEqpNm : ����
'       - pBarNo : ���ڵ��ȣ
'       - pTitle : ��������
'-----------------------------------------------------------------------------'
Public Function AddX(ByVal pCode As String, ByVal pEqpCd As String, ByVal pEqpNm As String, _
                     Optional pBarNo As String, Optional ByVal pTitle As String) As clsIISIntError
    Dim objIntErr As clsIISIntError
    Dim lngSeq    As Long       '��������
    
    lngSeq = mIntErrs.Count + 1
    Set objIntErr = New clsIISIntError
    With objIntErr
        .Seq = lngSeq
        .Code = pCode
        .ErrDt = Format$(Now, "YYYY-MM-DD HH:MM:SS")
        .EqpCd = pEqpCd
        .EqpNm = pEqpNm
        .BarNo = pBarNo
        .Title = pTitle
    End With
    
    mIntErrs.Add objIntErr, CStr(lngSeq)
    
    Set AddX = objIntErr
    Set objIntErr = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'   �μ� :
'       - pTestCd   : �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pSeq As Long)
    mIntErrs.Remove CStr(pSeq)
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mIntErrs.Count To 1 Step -1
        mIntErrs.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   �μ� :
'       - pSeq : ��������
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pSeq As Long) As Boolean
    Dim objIntErr As clsIISIntError
    
    For Each objIntErr In mIntErrs
        With objIntErr
            If pSeq = .Seq Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objIntErr = Nothing
End Function

Private Sub Class_Initialize()
    Set mIntErrs = New Collection
End Sub

Private Sub Class_Terminate()
    Set mIntErrs = Nothing
End Sub
