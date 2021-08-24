VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISResult.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : ������� Ŭ����
'   �ۼ���  : 2004-03-31
'   ��  ��  :
'       1. 1.1.4: �̻��(2005-04-11)
'       2. 1.1.5: �̻��(2005-04-12)
'       3. 1.1.11: �̻��(2005-07-22)
'       4. 1.1.12: �̻��(2005-07-28)
'       5. 1.1.14: �̻��(2005-10-11)
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISLibrarySql

Private mIntNm      As clsIISIntNm      '��� �˻��׸� Ŭ����
Private mTSpc       As clsIISTSpc       '������ü Ŭ����
Private mRef        As clsIISRef        '����ġ Ŭ����

Private mTestCd     As String   '�˻��ڵ�
Private mTestNm10   As String   '�˻��(10�ڸ�)
Private mTestNm     As String   '�˻��(��ü)
Private mIntResult  As String   '��񿡼� ������ ���
Private mResult     As String   '��񿡼� ������ ��� (IIS���� �����Ѱ��)
Private mRstVal     As Single   '���1(Numeric)
Private mRstCd      As String   '���2(Alphabetic)
Private mUnit       As String   '�������
Private mHLDiv      As String   'H/L
Private mDPDiv      As String   'D/P
Private mLastRst    As String   '�ֱٰ��
Private mRerunFg    As String   '��˿��� (0:��, 1:��)
Private mMfyFg      As String   '�������� (0:��, 1:��)
Private mTxtFg      As String   '�ؽ�Ʈ��� (0:��, 1:��)
Private mAutoFg     As String   '������ۿ��� (0:��, 1:��)
Private mRstType    As String   '������� (F:Free, N:Numeric, A:Alpha)
Private mRstDiv     As String   'Require/Alternative ���� (R:Require, A:Alternative, *:�󼼸��ڵ�)
Private mDetailFg   As String   '���׸����
Private mEqpCd      As String   '����ڵ�
Private mAutoVfy    As Boolean  '�ڵ������� ���ɿ���
Private mIntInfo    As String   '��񿡼� ���۵� �߰�����
Private mSndChk     As Boolean  '�������� ���� (BN100������ ����)

Public Property Get IntNm() As clsIISIntNm
    Set IntNm = mIntNm
End Property

Public Property Let IntNm(ByVal vData As clsIISIntNm)
    Set mIntNm = vData
End Property

Public Property Get TSpc() As clsIISTSpc
    Set TSpc = mTSpc
End Property

Public Property Set TSpc(ByVal vData As clsIISTSpc)
    Set mTSpc = vData
End Property

Public Property Get Ref() As clsIISRef
    Set Ref = mRef
End Property

Public Property Set Ref(ByVal vData As clsIISRef)
    Set mRef = vData
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

Public Property Get IntResult() As String
    IntResult = mIntResult
End Property

Public Property Let IntResult(ByVal vData As String)
    mIntResult = vData
End Property

Public Property Get Result() As String
    Result = mResult
End Property

Public Property Let Result(ByVal vData As String)
    mResult = vData
End Property

'## 1.1.4: �̻��(2005-04-11)
'   - mRstVal�� �ڷ����� Long���� Single�� ����
Public Property Get RstVal() As Single
    RstVal = mRstVal
End Property

Public Property Let RstVal(ByVal vData As Single)
    mRstVal = vData
End Property

Public Property Get RstCd() As String
    RstCd = mRstCd
End Property

Public Property Let RstCd(ByVal vData As String)
    mRstCd = vData
End Property

Public Property Get Unit() As String
    Unit = mUnit
End Property

Public Property Let Unit(ByVal vData As String)
    mUnit = vData
End Property

Public Property Get HLDiv() As String
    HLDiv = mHLDiv
End Property

Public Property Let HLDiv(ByVal vData As String)
    mHLDiv = vData
End Property

Public Property Get DPDiv() As String
    DPDiv = mDPDiv
End Property

Public Property Let DPDiv(ByVal vData As String)
    mDPDiv = vData
End Property

Public Property Get LastRst() As String
    LastRst = mLastRst
End Property

Public Property Let LastRst(ByVal vData As String)
    mLastRst = vData
End Property

Public Property Get RerunFg() As String
    RerunFg = mRerunFg
End Property

Public Property Let RerunFg(ByVal vData As String)
    mRerunFg = vData
End Property

Public Property Get MfyFg() As String
    MfyFg = mMfyFg
End Property

Public Property Let MfyFg(ByVal vData As String)
    mMfyFg = vData
End Property

Public Property Get TxtFg() As String
    TxtFg = mTxtFg
End Property

Public Property Let TxtFg(ByVal vData As String)
    mTxtFg = vData
End Property

Public Property Get AutoFg() As String
    AutoFg = mAutoFg
End Property

Public Property Let AutoFg(ByVal vData As String)
    mAutoFg = vData
End Property

Public Property Get RstType() As String
    RstType = mRstType
End Property

Public Property Let RstType(ByVal vData As String)
    mRstType = vData
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

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get AutoVfy() As Boolean
    AutoVfy = mAutoVfy
End Property

Public Property Let AutoVfy(ByVal vData As Boolean)
    mAutoVfy = vData
End Property

Public Property Get IntInfo() As String
    IntInfo = mIntInfo
End Property

Public Property Let IntInfo(ByVal vData As String)
    mIntInfo = vData
End Property

Public Property Get SndChk() As Boolean
    SndChk = mSndChk
End Property

Public Property Let SndChk(ByVal vData As Boolean)
    mSndChk = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ������ȣ�� �̿��� ������� ��ȸ
'   �μ� :
'       - pWorkarea  : Workarea
'       - pAccDt     : ��������
'       - pAccSeq    : ��������
'       - pAutoRerun : �ڵ���� �������
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pAutoRerun As AutoRerunEnum) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
    '## 1.1.12: �̻��(2005-07-28)
    '   - �ڵ���� ������� �μ��߰�
    '   - �ڵ������ ����Ҷ��� ����� ����־ �������� ���� �׸��� ��ȸ
On Error GoTo Errors:
    If pAutoRerun = ccAutoRerun Then
        Set Rs = DbCon.Execute(mSQL.SelectAutoRerunResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    Else
        Set Rs = DbCon.Execute(mSQL.SelectResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    End If
    Set GetResultInfo = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISResult", "GetResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������ȣ�� �̿��� Ư���˻� ������� ��ȸ - 1.1.14: �̻��(2005-10-11)
'          Ư���˻� ���Ŭ������ �����ؾ��ϴ°� ��Ģ������ ����� ������ �ʿ䰡
'          ���� �Ϲݰ˻� ���Ŭ������ ������
'   �μ� :
'       - pWorkarea  : Workarea
'       - pAccDt     : ��������
'       - pAccSeq    : ��������
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetSpecialResultInfo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
'    Set Rs = DbCon.Execute(mSQL.SelectSpecialResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    Set Rs = DbCon.Execute(mSQL.SelectResultInfo(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    Set GetSpecialResultInfo = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISResult", "GetSpecialResultInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻��� üũ
'-----------------------------------------------------------------------------'
Public Sub CheckResult()
    If mResult = "" Then Exit Sub
    
    '## ������� üũ: ��������� N, A, Null�ϰ�쿡�� ������� �̿��� üũ
    If mRstType = "A" Or mRstType = "N" Or mRstType = "" Then
        If Not (mResult Like "*+" Or mResult Like "*-") And IsNumeric(mResult) Then
            mRstType = "N"
        Else
            mRstType = "F"
        End If
    End If
    
    '## ��ȿ���� üũ
    Call CheckAvalValue
    
    '## 1.1.5: �̻��(2005-04-12)
    '   - ��������� Numeric�϶��� Delta, Panic, High, Low üũ
    '   - ��������� Free�϶��� Alpha��� ����ġ üũ �ϵ��� ����
    '## ����� Null�� �ƴҶ��� H/P, D/Püũ
    If mResult <> "" Then
        If mRstType = "N" Then
            Call CheckDeltaPanic
            Call CheckHighLow
        ElseIf mRstType = "F" Then
            Call CheckAlphaValue
        End If
    End If
    
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �˻����� ��ȿ���� üũ
'-----------------------------------------------------------------------------'
Public Sub CheckAvalValue()
    Dim strFormat  As String     '�Ҽ��� �ڸ����� ����
    Dim strResult  As String     '�ӽð��
    Dim lngAvalVal As Long       '��ȿ����

On Error Resume Next
    
    '## ��ȿ���� üũ
    If mRstType = "N" Then
        lngAvalVal = mTSpc.AvalVal
        If lngAvalVal > 0 And lngAvalVal < 9 Then
            strFormat = "0." & String$(lngAvalVal, "0")
            strResult = Format$(mResult, strFormat)
        ElseIf lngAvalVal = "0" Then
            strResult = Round(mResult)
        Else
            strResult = mResult
        End If
    End If
    
    '## ��������� ���� ����� �Է�
    Select Case mRstType
        Case "N"    '## Numeric
            mRstVal = CSng(strResult)
            mRstCd = strResult
        Case "F"    '## Free
            mRstCd = mResult
    End Select
    
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �˻����� Delta/Panic üũ
'   ��ȯ :
'       - D : Delta
'       - P : Panic
'       - DP : Delta & Panic
'-----------------------------------------------------------------------------'
Public Function CheckDeltaPanic() As String
    Dim strPanic As String  '�ӽ� Panic
    Dim strDelta As String  '�ӽ� Delta
    Dim lngTemp  As String

On Error Resume Next
    
    If IsNumeric(mResult) = False Then Exit Function
    
    '## Panic üũ
    If IISPANICCHECK = "IIS102" Then
        '## ������ü �����Ϳ��� Panic üũ
        If mTSpc.PanicFg = "1" Then
            If Val(mResult) < mTSpc.PanicFrVal Or Val(mResult) > mTSpc.PanicToVal Then
                strPanic = "P"
            End If
        End If
    ElseIf IISPANICCHECK = "IIS103" Then
        '## ����ġ �����Ϳ��� Panic üũ
        If mRef.PanicFg = "1" Then
            If Val(mResult) < mRef.PanicFrVal Or Val(mResult) > mRef.PanicToVal Then
                strPanic = "P"
            End If
        End If
    End If
    
    '## Delta üũ
    If mTSpc.DeltaFg = "1" And mLastRst <> "" And IsNumeric(mLastRst) Then
        '## 1.1.11: �̻��(2005-07-22)
        '   - ���Ű���� "0"�϶��� ó�� �߰�
        If Val(mLastRst) = 0 Then
            lngTemp = 0
        Else
            lngTemp = ((Val(mLastRst) - Val(mResult)) / Val(mLastRst)) * 100
        End If
        
        If lngTemp < mTSpc.DeltaFrVal * (-1) Or lngTemp > TSpc.DeltaToVal Then
            strDelta = "D"
        End If
    End If
    
    mDPDiv = strPanic & strDelta
    CheckDeltaPanic = mDPDiv
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻����� High/Low üũ
'   ��ȯ :
'       - L  : Low
'       - H  : High
'       - LH : Low & High
'-----------------------------------------------------------------------------'
Public Function CheckHighLow() As String
    If IsNumeric(mResult) = False Then Exit Function
    
    If mRef.RefFg = "1" Then
        If Val(mResult) < mRef.RefFrVal Then
            mHLDiv = "L"
        ElseIf Val(mResult) > mRef.RefToVal Then
            mHLDiv = "H"
        End If
    End If
    
    CheckHighLow = mHLDiv
End Function

'-----------------------------------------------------------------------------'
'   ��� : Alpha��� ����ġ üũ - 1.1.5: �̻��(2005-04-12)
'-----------------------------------------------------------------------------'
Private Sub CheckAlphaValue()
    Dim strDecision As String
    
    strDecision = mRef.GetAlpahRef(mTestCd, mRstCd)
    Select Case strDecision
        Case "H":   mHLDiv = "H"    'High
        Case "L":   mHLDiv = "L"    'Low
        Case "P":   mDPDiv = "P"    'Panic
        Case "*":   mDPDiv = "N"    'Abnormal
    End Select
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �������̽� ����Է½� �����Ǵ� ���� ��ȸ
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pEqpCd    : ����ڵ�
'   ��ȯ : ��������� ������Ʈ �ϴ�����
'-----------------------------------------------------------------------------'
Public Function GetIISRegisterSql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pEqpCd As String) As String
    Dim SQL As String
    
    '## �󼼸��ڵ� �϶�: AUTOFG, EQPCD
    '## �����׸� �϶�: RSTVAL, RSTCD, HLDIV, DPDIV, AUTOFG, RSTTYPE, EQPCD,
    If mRstDiv = "*" Then
        SQL = mSQL.UpdateIISResultDetail(pWorkarea, pAccDt, pAccSeq, mTestCd, "1", pEqpCd)
    Else
        If mRstType = "N" Then
            SQL = mSQL.UpdateIISResultNumeric(pWorkarea, pAccDt, pAccSeq, mTestCd, mRstVal, mRstCd, mHLDiv, _
                    mDPDiv, "1", mRstType, pEqpCd)
        Else
            SQL = mSQL.UpdateIISResultFree(pWorkarea, pAccDt, pAccSeq, mTestCd, mRstCd, mHLDiv, _
                    mDPDiv, "1", mRstType, pEqpCd)
        End If
    End If
    
    GetIISRegisterSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������̽� �ڵ������Ͻ� �����Ǵ� ���� ��ȸ
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'       - pEqpCd    : ����ڵ�
'       - pVfyDt    : ���Ȯ������
'       - pVfyTm    : ���Ȯ�νð�
'   ��ȯ : ��������� ������Ʈ �ϴ�����
'-----------------------------------------------------------------------------'
Public Function GetIISVfySql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pEqpCd As String, ByVal pVfyDt As String, _
                    ByVal pVfyTm As String) As String
    Dim SQL      As String
    
    If mRstDiv = "*" Then
        SQL = mSQL.UpdateIISVfyDetail(pWorkarea, pAccDt, pAccSeq, mTestCd, "1", pEqpCd, pVfyDt, _
                pVfyTm, EmpId)
    Else
        SQL = mSQL.UpdateIISVfy(pWorkarea, pAccDt, pAccSeq, mTestCd, mRstVal, mRstCd, mHLDiv, _
                mDPDiv, "1", mRstType, pEqpCd, pVfyDt, pVfyTm, EmpId)
    End If
    
    GetIISVfySql = SQL
End Function

Private Sub Class_Initialize()
    Set mIntNm = New clsIISIntNm
    Set mTSpc = New clsIISTSpc
    Set mRef = New clsIISRef
    Set mSQL = New clsIISLibrarySql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
    Set mRef = Nothing
    Set mTSpc = Nothing
    Set mIntNm = Nothing
End Sub