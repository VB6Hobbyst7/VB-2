VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISAccInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISAccInfos.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : �������� �÷��� Ŭ����
'   �ۼ���  : 2004-04-01
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfos As Collection

Public Property Get Count() As Long
    Count = mAccInfos.Count
End Property

Public Property Get Item(ByVal pSpcYy As String, ByVal pSpcNo As Long) As clsIISAccInfo
Attribute Item.VB_UserMemId = 0
    Set Item = mAccInfos(pSpcYy & DIV & CStr(pSpcNo))
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mAccInfos.[_NewEnum]
End Property
    
'-----------------------------------------------------------------------------'
'   ��� : �������� �߰�1 (�������̽���)
'   �μ� :
'       - pSpcYy      : ��ü����
'       - pSpcNo      : ��ü��ȣ
'       - pPtid       : ȯ��ID
'       - pName       : ȯ�ڸ�
'       - pSsn        : �ֹι�ȣ
'       - pAgeDay     : �Ϸ�
'       - pSex        : ����
'       - pDeptNm     : �������
'       - pOrdDoctNm  : ó���Ǹ�
'       - pWorkarea   : Workarea
'       - pAccDt      : ��������
'       - pAccSeq     : ��������
'       - pStsCd      : ��ü���� (0:ó��, 1:ä��, 2:����, 3: �˻���, 4:�߰����, 5:�Ϸ�, 6:����, 7:�������)
'       - pReqTotCnt  : �� �˻��׸� ����
'       - pInputCnt   : �Էµ� �˻��׸� ����
'       - pSpcCd      : ��ü�ڵ�
'       - pSpcNm      : ��ü��
'       - pWardNm     : ������
'       - pTestDiv    : �Ϲ�, �̻���, ��Ÿ�˻� ���� (0:�Ϲݰ˻�, 1:��Ÿ�˻�, 2:�̻����˻�)
'       - pQcFg       : QC ���� (0:�Ϲ�ó��, 1:QCó��)
'       - pStatFg     : ���޿��� (0:������, 1:����)
'       - pBuildCd    : �ǹ��ڵ�
'       - pHosilId    : ȣ���ڵ�
'   ��ȯ : �������� Ŭ����
'-----------------------------------------------------------------------------'
Public Function AddInter(ByVal pSpcYy As String, ByVal pSpcNo As Long, _
                    Optional ByVal pPtId As String, Optional ByVal pName As String, _
                    Optional ByVal pSsn As String, Optional ByVal pAgeDay As String, _
                    Optional ByVal pSex As String, Optional ByVal pDeptNm As String, _
                    Optional ByVal pOrdDoctNm As String, Optional ByVal pWorkarea As String, _
                    Optional ByVal pAccDt As String, Optional ByVal pAccSeq As String, _
                    Optional ByVal pStsCd As String, Optional ByVal pReqTotCnt As Long, _
                    Optional ByVal pInputCnt As Long, Optional ByVal pSpcCd As String, _
                    Optional ByVal pSpcNm As String, Optional ByVal pWardNm As String, _
                    Optional ByVal pTestDiv As String, Optional ByVal pQcFg As String, _
                    Optional ByVal pStatFg As String, Optional ByVal pBuildCd As String, _
                    Optional ByVal pHosilId As String) As clsIISAccInfo
    Dim objAccInfo As clsIISAccInfo
    
    Set objAccInfo = New clsIISAccInfo
    With objAccInfo
        .SpcYy = pSpcYy:            .SpcNo = pSpcNo
        .PtId = pPtId:              .Name = pName
        .Ssn = pSsn:                .AgeDay = pAgeDay
        .Sex = pSex:                .DeptNm = pDeptNm
        .OrdDoctNm = pOrdDoctNm:    .Workarea = pWorkarea
        .AccDt = pAccDt:            .AccSeq = pAccSeq
        .StsCd = pStsCd:            .ReqTotCnt = pReqTotCnt
        .InputCnt = pInputCnt:      .SpcCd = pSpcCd
        .SpcNm = pSpcNm:            .WardNm = pWardNm
        .TestDiv = pTestDiv:        .QcFg = pQcFg
        .StatFg = pStatFg:          .BuildCd = pBuildCd
        .HosilId = pHosilId
    End With
    
    mAccInfos.Add objAccInfo, pSpcYy & DIV & CStr(pSpcNo)
    
    Set AddInter = objAccInfo
    Set objAccInfo = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� �߰�1 (�������̽���)
'   �μ� :
'       - pAccInfo : �������� ��ü
'-----------------------------------------------------------------------------'
Public Sub AddInterX(ByVal pAccInfo As clsIISAccInfo)
    mAccInfos.Add pAccInfo, pAccInfo.SpcYy & DIV & CStr(pAccInfo.SpcNo)
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'   �μ� :
'       - pSpcYy : ��ü����
'       - pSpcNo : ��ü��ȣ
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pSpcYy As String, ByVal pSpcNo As Long)
    mAccInfos.Remove pSpcYy & DIV & CStr(pSpcNo)
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mAccInfos.Count To 1 Step -1
        mAccInfos.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Index�� ���������� ��ȯ
'   �μ� :
'       - pSpcYy : ��ü����
'       - pSpcNo : ��ü��ȣ
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pSpcYy As String, ByVal pSpcNo As Long) As Boolean
    Dim objAccInfo As clsIISAccInfo
    
    For Each objAccInfo In mAccInfos
        With objAccInfo
            If pSpcYy = .SpcYy And pSpcNo = .SpcNo Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objAccInfo = Nothing
End Function

Private Sub Class_Initialize()
    Set mAccInfos = New Collection
End Sub

Private Sub Class_Terminate()
    Set mAccInfos = Nothing
End Sub
