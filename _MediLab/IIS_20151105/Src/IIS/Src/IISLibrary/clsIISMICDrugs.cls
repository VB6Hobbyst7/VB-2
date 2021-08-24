VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMICDrugs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsMICIISDrugs.cls
'   �ۼ���  : �̻��
'   ��  ��  : Vitek ��� �׻������ �÷���
'   �ۼ���  : 2004-09-23
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mDrugs As Collection

Public Property Get Count() As Long
    Count = mDrugs.Count
End Property

Public Property Get Item(ByVal pDrugCd As String) As clsIISMICDrug
Attribute Item.VB_UserMemId = 0
    Set Item = mDrugs(pDrugCd)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mDrugs.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ� �׻������ �߰�
'   �μ� :
'       1.pDrugCd : �׻����ڵ�
'       2.pDrugNm : �׻�����
'       3.pVolumn : �Է�
'       4.pRstCd  : ����ڵ�
'   ��ȯ : �׻������ Ŭ����'
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pDrugCd As String, ByVal pDrugNm As String, ByVal pVolumn As String, _
                    ByVal pRstCd As String) As clsIISMICDrug
    Dim objDrug As clsIISMICDrug
    
    Set objDrug = New clsIISMICDrug
    With objDrug
        .DrugCd = pDrugCd
        .DrugNm = pDrugNm
        .Volumn = pVolumn
        .RstCd = pRstCd
    End With
    
    mDrugs.Add objDrug, pDrugCd
    
    Set Add = objDrug
    Set objDrug = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pDrugCd As String)
    mDrugs.Remove pDrugCd
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mDrugs.Count To 1 Step -1
        mDrugs.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pDrugCd As String) As Boolean
    Dim objDrug As clsIISMICDrug
    
    For Each objDrug In mDrugs
        If objDrug.DrugCd = pDrugCd Then
            Exist = True
            GoTo EndLine
        End If
    Next
    Exist = False

EndLine:
    Set objDrug = Nothing
End Function

Private Sub Class_Initialize()
    Set mDrugs = New Collection
End Sub

Private Sub Class_Terminate()
    Set mDrugs = Nothing
End Sub
