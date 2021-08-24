VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMICResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISMICResult.cls
'   �ۼ���  : �̻��
'   ��  ��  : �̻��� ������� Ŭ����
'   �ۼ���  : 20034-09-08
'   ��  ��  :
'       1. 1.0.2: �̻��(2004-12-06)
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMicroSql
Private mWSBody     As clsIISMICWSBody      '�̻��� WorkSheet Body Ŭ����
Private mSensi      As clsIISMICSensi       '�̻��� ��������� Ŭ����

Private mTestCd     As String   '�˻��ڵ�
Private mTestNm10   As String   '�˻��(10�ڸ�)
Private mTestNm     As String   '�˻��(Full)
Private mRstCd      As String   '���
Private mSenFg      As String   '������ �������(Y:��, Null:��)
Private mStsCd      As String   '��ü����(0:ó��, 1:ä��, 2:����, 3:�˻���, 4:�߰����, 5:���, 6:����, 7:�������)
Private mRstType    As String   '�������(S:�Ϲݰ�����, M:MIC������, G:Gram Stain, M:AFB Culture, F:Fungus Culture, N:Fungus Stain, B:AFB Stain, A:AFB Sensitivity)
Private mMfySeq     As Long     '����ȸ��
Private mRstDiv     As String   '�������(A:Alternative, R:Require, *:�������)
Private mDetailFg   As String   '���׸����
Private mVfyDt      As String   '���Ȯ������
Private mVfyTm      As String   '���Ȯ�νð�
Private mVfyId      As String   '���Ȯ����ID
Private mPtId       As String   'ȯ��ID

Public Property Get WSBody() As clsIISMICWSBody
    Set WSBody = mWSBody
End Property

Public Property Let WSBody(ByVal vData As clsIISMICWSBody)
    Set mWSBody = vData
End Property

Public Property Get Sensi() As clsIISMICSensi
    Set Sensi = mSensi
End Property

Public Property Let Sensi(ByVal vData As clsIISMICSensi)
    Set mSensi = vData
End Property

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get TestNm10() As String
    TestNm10 = mTestNm10
End Property

Public Property Let TestNm10(ByVal vData As String)
    mTestNm10 = vData
End Property

Public Property Get TestNm() As String
    TestNm = mTestNm
End Property

Public Property Let TestNm(ByVal vData As String)
    mTestNm = vData
End Property

Public Property Get RstCd() As String
    RstCd = mRstCd
End Property

Public Property Let RstCd(ByVal vData As String)
    mRstCd = vData
End Property

Public Property Get SenFg() As String
    SenFg = mSenFg
End Property

Public Property Let SenFg(ByVal vData As String)
    mSenFg = vData
End Property

Public Property Get StsCd() As String
    StsCd = mStsCd
End Property

Public Property Let StsCd(ByVal vData As String)
    mStsCd = vData
End Property

Public Property Get RstType() As String
    RstType = mRstType
End Property

Public Property Let RstType(ByVal vData As String)
    mRstType = vData
End Property

Public Property Get MfySeq() As Long
    MfySeq = mMfySeq
End Property

Public Property Let MfySeq(ByVal vData As Long)
    mMfySeq = vData
End Property

Public Property Get RstDiv() As String
    RstDiv = mRstDiv
End Property

Public Property Let RstDiv(ByVal vData As String)
    mRstDiv = vData
End Property

Public Property Get DetailFg() As String
    DetailFg = mDetailFg
End Property

Public Property Let DetailFg(ByVal vData As String)
    mDetailFg = vData
End Property

Public Property Get VfyDt() As String
    VfyDt = mVfyDt
End Property

Public Property Let VfyDt(ByVal vData As String)
    mVfyDt = vData
End Property

Public Property Get VfyTm() As String
    VfyTm = mVfyTm
End Property

Public Property Let VfyTm(ByVal vData As String)
    mVfyTm = vData
End Property

Public Property Get VfyId() As String
    VfyId = mVfyId
End Property

Public Property Let VfyId(ByVal vData As String)
    mVfyId = vData
End Property

Public Property Get Ptid() As String
    Ptid = mPtId
End Property

Public Property Let Ptid(ByVal vData As String)
    mPtId = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ������ȣ�� �̿��� �̻��� ������� ��ȸ
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pTestCds  : ��� �˻��׸� ��ϵ� �˻��ڵ�
'   ��ȯ :
'       -1 : ��������� ����
'-----------------------------------------------------------------------------'
Public Function GetMICResultInfo(ByVal pWorkArea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCds As String) As Long
    Dim rs As ADODB.Recordset
    
    '## 1.0.2: �̻��(2004-12-06)
    '   - ��� �˻��׸� ��ϵ� �ڵ常 ��ȸ�ϵ��� pTestCds �߰�
On Error GoTo Errors
    Set rs = DbCon.Execute(mSQL.SelectMICResultInfo(pWorkArea, pAccDt, pAccSeq, pTestCds), , adCmdText)
    If Not (rs.BOF Or rs.EOF) Then
        mTestCd = rs.Fields("TESTCD").Value
        mTestNm10 = rs.Fields("TESTNM10").Value & ""
        mTestNm = rs.Fields("TESTNM").Value & ""
        mRstType = rs.Fields("RSTTYPE").Value & ""
        mMfySeq = CLng(rs.Fields("MFYSEQ").Value)
        mPtId = rs.Fields("PTID").Value & ""
        
        With mWSBody
            .WsCd = rs.Fields("WSCD").Value
            .WsUnit = rs.Fields("WSUNIT").Value
            .SelFg = rs.Fields("SELFG").Value & ""
            .GenFg = rs.Fields("GENFG").Value & ""
            .ScFg = rs.Fields("SCFG").Value & ""
        End With
        GetMICResultInfo = 1
    Else
        GetMICResultInfo = -1
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Function
    
Errors:
    Set rs = Nothing
    GetMICResultInfo = -1
    Error.SetLog App.EXEName, "clsIISMICResult", "GetMICResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

''-----------------------------------------------------------------------------'
''   ��� : ������ȣ�� �̿��� �̻��� ������� ��ȸ
''   �μ� :
''       - pWorkarea : Workarea
''       - pAccDt    : ��������
''       - pAccSeq   : ��������
''   ��ȯ :
''       -1 : ��������� ����
''-----------------------------------------------------------------------------'
'Public Function GetCultureResultInfo(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As Long
'    Dim rs As ADODB.Recordset
'
'On Error GoTo Errors
'    Set rs = DbCon.Execute(mSQL.SelectCultureResultInfo(pWorkArea, pAccDt, pAccSeq), , adCmdText)
'    If Not (rs.BOF Or rs.EOF) Then
'        mStsCd = rs.Fields("STSCD").Value & ""
'        mTestCd = rs.Fields("TESTCD").Value & ""
'
'        GetCultureResultInfo = 1
'    Else
'        GetCultureResultInfo = -1
'    End If
'
'    rs.Close
'    Set rs = Nothing
'    Exit Function
'
'Errors:
'    Set rs = Nothing
'    GetCultureResultInfo = -1
'    Error.SetLog App.EXEName, "clsIISMICResult", "GetCultureResultInfo", Err.Description, Now
'    MsgBox Error.Description, vbCritical, "����"
'End Function
'
''-----------------------------------------------------------------------------'
''   ��� : �̻������ �������� ��ȸ
''   �μ� :
''       - pWorkarea : Workarea
''       - pAccDt    : ��������
''       - pAccSeq   : ��������
''-----------------------------------------------------------------------------'
'Public Function SaveMicroResultSql(ByVal pWorkArea As String, ByVal pAccDt As String, _
'                    ByVal pAccSeq As Long) As String
'
'    mSenFg = "Y"
'    SaveMicroResultSql = mSQL.UpdateMicroResult(pWorkArea, pAccDt, pAccSeq, mSenFg)
'End Function
'
''-----------------------------------------------------------------------------'
''   ��� : �̻��� ����� �������� ��ȸ
''   �μ� :
''       - pWorkarea : Workarea
''       - pAccDt    : ��������
''       - pAccSeq   : ��������
''-----------------------------------------------------------------------------'
'Public Function SaveCultureResultSql(ByVal pWorkArea As String, ByVal pAccDt As String, _
'                    ByVal pAccSeq As Long, ByVal pIntInfo As clsIISIntInfo) As String
'
'    Dim strResult As String
'
'    '-- ���
'    strResult = pIntInfo.VitekNo & "|" & mTestCd
'    If pIntInfo.VitekNo <> "C12" Then
'        SaveCultureResultSql = mSQL.UpdateCultureResult(pWorkArea, pAccDt, pAccSeq, strResult)
'    Else
'        SaveCultureResultSql = ""
'    End If
'
'End Function
'
''-----------------------------------------------------------------------------'
''   ��� : �̻��� ��� �������� �������� ��ȸ
''   �μ� :
''       - pWorkarea : Workarea
''       - pAccDt    : ��������
''       - pAccSeq   : ��������
''-----------------------------------------------------------------------------'
'Public Function SaveCultureInfoSql(ByVal pWorkArea As String, ByVal pAccDt As String, _
'                    ByVal pAccSeq As Long, ByVal pIntInfo As clsIISIntInfo) As String
'
'
'    Dim strResult As String
'
'    '-- ���
'    strResult = pIntInfo.VitekNo
'    If strResult <> "C12" Then  '�缺�� ������Ʈ ���� �ʴ´�.
'        SaveCultureInfoSql = mSQL.UpdateCultureInfoResult(pWorkArea, pAccDt, pAccSeq, strResult)
'    Else
'        SaveCultureInfoSql = ""
'    End If
'
'End Function
''-----------------------------------------------------------------------------'
''   ��� : �̻��� ��������� �������� ��ȸ
''   �μ� :
''       - pWorkarea : Workarea
''       - pAccDt    : ��������
''       - pAccSeq   : ��������
''       - pIntInfo  : ��񿡼� �����Ѱ�� ����Ŭ����
''-----------------------------------------------------------------------------'
'Public Function SaveMICResultSql(ByVal pWorkArea As String, ByVal pAccDt As String, _
'                    ByVal pAccSeq As Long, ByVal pIntInfo As clsIISIntInfo) As String
'    Dim objDrug   As clsIISMICDrug  '�׻������ Ŭ����
'    Dim strResult As String         '�׻������
'    Dim SQL       As String
'
'    '## NOTE: ���ڿ� �ݵ�� �����Ұ�!
'    '## �׻�������� DB�� �Է��������·� ����
'    For Each objDrug In pIntInfo.Drugs
'        With objDrug
'            strResult = strResult & ", '" & .DrugCd & ";" & .RstCd & ";" & .Volumn & ";;;;'"
'        End With
'    Next
'    Set objDrug = Nothing
'
'    With mSensi
'        .MfySeq = 0
'        .MicFg = "C"
'        .MnmCd = pIntInfo.MnmCd
'        .MicFg = "C"
'        .MqtCd = IISMQTCD
'        .SCnt = pIntInfo.Drugs.Count
'        .Result = strResult
'        SQL = .GetSaveMICResultSql(pWorkArea, pAccDt, pAccSeq, mTestCd)
'    End With
'
'    SaveMICResultSql = SQL
'End Function



Private Sub Class_Initialize()
    Set mWSBody = New clsIISMICWSBody
    Set mSensi = New clsIISMICSensi
    Set mSQL = New clsIISMicroSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
    Set mSensi = Nothing
    Set mWSBody = Nothing
End Sub