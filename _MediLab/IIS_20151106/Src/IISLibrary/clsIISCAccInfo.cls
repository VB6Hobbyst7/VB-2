VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCAccInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISCAccInfo.cls
'   �ۼ���  : �̻��
'   ��  ��  : ClientDb�� �������� Ŭ����
'   �ۼ���  : 2004-04-29
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISLibrarySql

Private mItemSeq    As Long     '���ۼ���
Private mSpcNo      As String   '��ü��ȣ(���ڵ��ȣ)
Private mPtId       As String   'ȯ��ID
Private mName       As String   '�̸�
Private mSex        As String   '����
Private mAgeDay     As Long     '�Ϸ�
Private mDoctNm     As String   'ó���Ǹ�
Private mDeptNm     As String   '�������
Private mWardNm     As String   '������
Private mSpcCd      As String   '��ü�ڵ�
Private mSpcNm      As String   '��ü��
Private mStatFg     As String   '���޿��� (0:������, 1:����)
Private mQcFg       As String   'QC���� (0:��, 1:��)
Private mCtrlCd     As String   'Control Code
Private mLevelCd    As String   'Level Code
Private mLotNo      As String   'Lot No
Private mSpcPos     As String   '��񿡼� ��ü��ġ
Private mEqpCd      As String   '����ڵ�
Private mTransDt    As String   '��������
Private mTransTm    As String   '�����Ͻ�

Public Property Get ItemSeq() As Long
    ItemSeq = mItemSeq
End Property

Public Property Let ItemSeq(ByVal vData As Long)
    mItemSeq = vData
End Property

Public Property Get SpcNo() As String
    SpcNo = mSpcNo
End Property

Public Property Let SpcNo(ByVal vData As String)
    mSpcNo = vData
End Property

Public Property Get PtId() As String
    PtId = mPtId
End Property

Public Property Let PtId(ByVal vData As String)
    mPtId = vData
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal vData As String)
    mName = vData
End Property

Public Property Get Sex() As String
    Sex = mSex
End Property

Public Property Let Sex(ByVal vData As String)
    mSex = vData
End Property

Public Property Get AgeDay() As Long
    AgeDay = mAgeDay
End Property

Public Property Let AgeDay(ByVal vData As Long)
    mAgeDay = vData
End Property

Public Property Get DoctNm() As String
    DoctNm = mDoctNm
End Property

Public Property Let DoctNm(ByVal vData As String)
    mDoctNm = vData
End Property

Public Property Get DeptNm() As String
    DeptNm = mDeptNm
End Property

Public Property Let DeptNm(ByVal vData As String)
    mDeptNm = vData
End Property

Public Property Get WardNm() As String
    WardNm = mWardNm
End Property

Public Property Let WardNm(ByVal vData As String)
    mWardNm = vData
End Property

Public Property Get SpcCd() As String
    SpcCd = mSpcCd
End Property

Public Property Let SpcCd(ByVal vData As String)
    mSpcCd = vData
End Property

Public Property Get SpcNm() As String
    SpcNm = mSpcNm
End Property

Public Property Let SpcNm(ByVal vData As String)
    mSpcNm = vData
End Property

Public Property Get StatFg() As String
    StatFg = mStatFg
End Property

Public Property Let StatFg(ByVal vData As String)
    mStatFg = vData
End Property

Public Property Get QcFg() As String
    QcFg = mQcFg
End Property

Public Property Let QcFg(ByVal vData As String)
    mQcFg = vData
End Property

Public Property Get SpcPos() As String
    SpcPos = mSpcPos
End Property

Public Property Let SpcPos(ByVal vData As String)
    mSpcPos = vData
End Property

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get TransDt() As String
    TransDt = mTransDt
End Property

Public Property Let TransDt(ByVal vData As String)
    mTransDt = vData
End Property

Public Property Get TransTm() As String
    TransTm = mTransTm
End Property

Public Property Let TransTm(ByVal vData As String)
    mTransTm = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ��ü�� �̿��� ClientDb�� �������
'   �μ� :
'       - pAccInfo : �������� ��ü
'       - pEqpCd   : ����ڵ�
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function SaveClientDb(ByVal pAccInfo As clsIISAccInfo, ByVal pEqpCd) As Boolean
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim objCResult  As clsIISCResult    'ClientDb ������� Ŭ����
    Dim arySQL()    As String           '����Է¿� ���Ǵ� SQL �迭
    Dim i           As Long
    
    With pAccInfo
        mItemSeq = GetMaxItemSeq
        mSpcNo = .GetBarNo:     mPtId = .PtId
        mName = .Name:          mSex = .Sex
        mAgeDay = .AgeDay:      mDoctNm = .OrdDoctNm
        mDeptNm = .DeptNm:      mWardNm = .WardNm
        mSpcCd = .SpcCd:        mSpcNm = Replace$(.SpcNm, "'", "''")
        mStatFg = .StatFg:      mQcFg = IIf(.QcFg = "1", "1", "0")
        mCtrlCd = .CtrlCd:      mLevelCd = .LevelCd
        mLotNo = .LotNo:        mSpcPos = .SpcPos
        mEqpCd = pEqpCd:        mTransDt = Format$(Now, "YYYYMMDD")
        mTransTm = Format$(Time, "HHMMSS")

        Set objCResult = New clsIISCResult
        If mQcFg = "0" Then         '## �Ϲݰ�ü
            '## 1.ClientDb�� �Ϲݰ˻� ���������� �Է��ϴ� ��������
            ReDim arySQL(1)
            arySQL(0) = GetCliRegisterSql
            
            '## 2.ClientDb�� �Ϲݰ˻� ��������� �Է��ϴ� ��������
            i = 1
            ReDim Preserve arySQL(.Results.Count + 1)
            For Each objResult In .Results
                If objResult.Result <> "" Then
                    arySQL(i) = objCResult.GetCliRegisterSql(mItemSeq, objResult.IntNm.IntNm, _
                        objResult.TestCd, objResult.Result, objResult.HLDiv, objResult.DPDiv)
                    i = i + 1
                End If
            Next
        ElseIf mQcFg = "1" Then     '## QC��ü
            '## 1.ClientDb�� QC�˻� ���������� �Է��ϴ� ��������
            ReDim arySQL(1)
            arySQL(0) = GetCliQCRegisterSql
            
            '## 2.ClientDb�� QC�˻� ��������� �Է��ϴ� ��������
            i = 1
            ReDim Preserve arySQL(.QCResults.Count + 1)
            For Each objQCResult In .QCResults
                If objQCResult.Result <> "" Then
                    arySQL(i) = objCResult.GetCliQCRegisterSql(mItemSeq, objQCResult.IntNm.IntNm, _
                        objQCResult.TestCd, objQCResult.Result, objQCResult.RADiv, objQCResult.Text)
                    i = i + 1
                End If
            Next
        End If
        
        Set objCResult = Nothing
        Set objQCResult = Nothing
        Set objResult = Nothing
    End With
    
On Error GoTo Errors
    CliCon.BeginTrans
    
    '## 3.ClientDb�� ����
    For i = 0 To UBound(arySQL)
        If arySQL(i) = "" Then Exit For
        CliCon.Execute arySQL(i), , adCmdText + adExecuteNoRecords
    Next i
    
    CliCon.CommitTrans
    SaveClientDb = True
    Exit Function

Errors:
    CliCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISCAccInfo", "SaveClientDb", Err.Description, Now
    Error.SetLog App.EXEName, "clsIISCAccInfo", "SaveClientDb", "SQL: " & arySQL(i), Now
End Function

'-----------------------------------------------------------------------------'
'   ��� : MAX(���ۼ���)+1 �� ��ȸ
'   ��ȯ : MAX(���ۼ���)+1
'-----------------------------------------------------------------------------'
Public Function GetMaxItemSeq() As Long
    Dim Rs As ADODB.Recordset
    
    Set Rs = CliCon.Execute(mSQL.SelectMaxSeq, , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        If IsNull(Rs.Fields("MAXSEQ").Value) Then
            GetMaxItemSeq = 1
        Else
            GetMaxItemSeq = CLng(Rs.Fields("MAXSEQ").Value)
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISCAccInfo", "GetMaxItemSeq", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb�� �Ϲݰ˻� ���������� �Է��ϴ� ������ȸ
'-----------------------------------------------------------------------------'
Public Function GetCliRegisterSql() As String
    Dim SQL As String
    
    SQL = mSQL.InsertClientAccInfo(mItemSeq, mSpcNo, mPtId, mName, mSex, mAgeDay, mDoctNm, mDeptNm, _
            mWardNm, mSpcCd, mSpcNm, mStatFg, mQcFg, mSpcPos, mEqpCd, mTransDt, mTransTm)
            
    GetCliRegisterSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb�� QC�˻� ���������� �Է��ϴ� ������ȸ
'-----------------------------------------------------------------------------'
Public Function GetCliQCRegisterSql() As String
    Dim SQL As String
    
    SQL = mSQL.InsertClientQCAccInfo(mItemSeq, mSpcNo, mPtId, mQcFg, mCtrlCd, mLevelCd, mLotNo, _
            mSpcPos, mEqpCd, mTransDt, mTransTm)
    
    GetCliQCRegisterSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb �������� ���� ������ ����
'   �μ� :
'       - pEqpCd   : ����ڵ�
'       - pStoreDt : ������
'-----------------------------------------------------------------------------'
Public Function DelHisClientDb(ByVal pEqpCd As String, ByVal pStoreDt As Long) As Boolean
    Dim strBaseDt As String     '����������
    Dim SQL       As String
    
    strBaseDt = Format$(DateAdd("d", -pStoreDt, Now), "YYYYMMDD")
    
On Error GoTo Errors
    '## ACC204 ���̺��� �����ʹ� Access���� Cascade�������� ������
    CliCon.BeginTrans
    CliCon.Execute mSQL.DeleteHisData(pEqpCd, strBaseDt), , adCmdText + adExecuteNoRecords
    CliCon.CommitTrans
    DelHisClientDb = True
    Exit Function

Errors:
    CliCon.RollbackTrans
    Error.SetLog App.EXEName, "clsIISCAccInfo", "DelHisClientDb", Err.Description, Now
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISLibrarySql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub