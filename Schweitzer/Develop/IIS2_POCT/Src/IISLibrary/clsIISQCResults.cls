VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISQCResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISQCResults.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : QC ������� �÷��� Ŭ����
'   �ۼ���  : 2004-05-21
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mQCResults As Collection

Public Property Get Count() As Long
    Count = mQCResults.Count
End Property

Public Property Get Item(ByVal pTestCd As String) As clsIISQCResult
Attribute Item.VB_UserMemId = 0
    Set Item = mQCResults(pTestCd)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mQCResults.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : QC������� �߰�1 (�������̽���)
'   �μ� :
'       - pTestCd   : �˻��ڵ�
'       - pTestNm10 : �˻��(10�ڸ�)
'       - pTestNm   : �˻��(��ü)
'       - pUnit     : �������
'       - pRstDiv   : Require/Alternative ���� (R:Require, A:Alternative, *:�󼼸��ڵ�)
'       - pCtrlCd   : Control Code
'       - pLevelCd  : Level Code
'       - pLotNo    : Lot No
'       - pQCMaster : QC Master ��ü
'   ��ȯ : QC�ᰡ���� ��ü
'-----------------------------------------------------------------------------'
Public Function AddIIS(ByVal pTestCd As String, Optional ByVal pTestNm10 As String, _
                    Optional ByVal pTestNm As String, Optional ByVal pUnit As String, _
                    Optional ByVal pRstDiv As String, Optional ByVal pCtrlCd As String, _
                    Optional ByVal pLevelCd As String, Optional ByVal pLotNo As String, _
                    Optional ByVal pQCMaster As clsIISQCMaster) As clsIISQCResult
    Dim objQCResult As clsIISQCResult
    
    Set objQCResult = New clsIISQCResult
    With objQCResult
        .TestCd = pTestCd
        .TestNm10 = pTestNm10
        .TestNm = pTestNm
        .Unit = pUnit
        .RstDiv = pRstDiv
        .CtrlCd = pCtrlCd
        .LevelCd = pLevelCd
        .LotNo = pLotNo
        .QCMaster = pQCMaster
    End With
    
    mQCResults.Add objQCResult, pTestCd
    
    Set AddIIS = objQCResult
    Set objQCResult = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : QC������� �߰�2 (�������̽���)
'   �μ� :
'       - pQCResult : QC������� ��ü
'   ��ȯ : QC������� ��ü
'-----------------------------------------------------------------------------'
Public Sub AddIISX(ByVal pQCResult As clsIISQCResult)
    mQCResults.Add pQCResult, pQCResult.TestCd
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'   �μ� :
'       - pSeq : ���ۼ���
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pTestCd As String)
    mQCResults.Remove pTestCd
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mQCResults.Count To 1 Step -1
        mQCResults.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   �μ� :
'       - pSeq : ���ۼ���
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    Dim objQCResult As clsIISQCResult
    
    For Each objQCResult In mQCResults
        With objQCResult
            If pTestCd = .TestCd Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objQCResult = Nothing
End Function

Private Sub Class_Initialize()
    Set mQCResults = New Collection
End Sub

Private Sub Class_Terminate()
    Set mQCResults = Nothing
End Sub
