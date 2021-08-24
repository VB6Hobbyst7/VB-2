VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISTMaster.cls
'   �ۼ���  : �̻��
'   ��  ��  : �˻��ڵ� ������ Ŭ����
'   �ۼ���  : 2004-01-19
'   ��  ��  :
'   ��  ��  : NOTE Ȯ���Ұ�!!
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mTestCds    As clsIISTestCds        '�˻��ڵ� �÷���
Private mFTestCds   As clsIISTestCdFulls    '�˻��ڵ�(��ü) �÷���
Private mTSpcs      As clsIISTSpcs          '������ü �÷���
Private mRefs       As clsIISRefs           '����ġ �÷���

Public Property Get TestCds() As clsIISTestCds
    Set TestCds = mTestCds
End Property

Public Property Let TestCds(ByRef vData As clsIISTestCds)
    Set mTestCds = vData
End Property

Public Property Get FTestCds() As clsIISTestCdFulls
    Set FTestCds = mFTestCds
End Property

Public Property Let FTestCds(ByRef vData As clsIISTestCdFulls)
    Set mFTestCds = vData
End Property

Public Property Get TSpcs() As clsIISTSpcs
    Set TSpcs = mTSpcs
End Property

Public Property Let TSpcs(ByRef vData As clsIISTSpcs)
    Set mTSpcs = vData
End Property

Public Property Get Refs() As clsIISRefs
    Set Refs = mRefs
End Property

Public Property Let Refs(ByRef vData As clsIISRefs)
    Set mRefs = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �ش� �˻��ڵ��� ���� �˻��ڵ带 ��ȸ
'          ���� �˻��ڵ尡 ������ ���� �˻��ڵ� ����
'   �μ� :
'       1.pTestCd : �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Function NextTestCd(ByVal pTestCd As String) As String
    Dim lngkey As Long      'Key
    
    lngkey = mTestCds(pTestCd).TestSeq + 1
    If lngkey > mTestCds.Count Then
        NextTestCd = pTestCd
        Exit Function
    End If
    
    NextTestCd = mTestCds(lngkey).TestCd
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� �˻��ڵ��� ���� �˻��ڵ带 ��ȸ
'          ���� �˻��ڵ尡 ������ ���� �˻��ڵ� ����
'   �μ� :
'       1.pTestCd : �˻��ڵ�
'-----------------------------------------------------------------------------'
Public Function PrevTestCd(ByVal pTestCd As String) As String
    Dim lngkey As Long      'Key
    
    lngkey = mTestCds(pTestCd).TestSeq - 1
    If lngkey <= 0 Then
        PrevTestCd = pTestCd
        Exit Function
    End If
    
    PrevTestCd = mTestCds(lngkey).TestCd
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� �˻��ڵ��� ���������� ��ȸ
'   �μ� :
'       1.pTestCd : �˻��ڵ�
'   ��ȯ : True(����), False(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    If mTestCds.Exist(pTestCd) Then
        Exist = True
    Else
        Exist = False
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ��ü�ڵ��� ���������� ��ȸ
'   �μ� :
'       1.pSpcCd  : ��ü�ڵ�
'   ��ȯ : True(����), False(������)
'-----------------------------------------------------------------------------'
Public Function ExistSpcCd(ByVal pSpcCd As String) As Boolean
    Dim objTSpc As clsIISTSpc   '������ü Ŭ����
    
    If mTSpcs.Count = 0 Then Exit Function
    For Each objTSpc In mTSpcs
        If pSpcCd = objTSpc.SpcCd Then
            ExistSpcCd = True
            Exit For
        End If
    Next
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ����ġ�� ���������� ��ȸ
'   �μ� :
'       1.pSpcCd  : ��ü�ڵ�
'   ��ȯ : True(����), False(������)
'-----------------------------------------------------------------------------'
Public Function ExistRef(ByVal pSpcCd As String) As Boolean
    Dim objRef As clsIISRef
    
    If mRefs.Count = 0 Then Exit Function
    For Each objRef In mRefs
        If pSpcCd = objRef.SpcCd Then
            ExistRef = True
            Exit For
        End If
    Next
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� �˻��ڵ��� �˻���� ��ȸ
'   �μ� :
'       1.pTestCd : �˻��ڵ�
'   ��ȯ : �˻��
'-----------------------------------------------------------------------------'
Public Function GetTestNm(ByVal pTestNm As String) As String
    '## �˻��ڵ尡 �ݵ�� �����ؾ� �Ѵ�.
    GetTestNm = mTestCds(pTestNm).TestNm
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� �˻��ڵ��� �����ֱ� �������� ��ȸ
'   �μ� :
'       1.mTestCd : �˻��ڵ�
'   ��ȯ :
'       1.�������� �����ϸ� ���� �ֱ� �������� ��ȯ
'       2.�������� ������ Null�� ��ȯ
'-----------------------------------------------------------------------------'
Public Function GetTestCdLastApplyDt(ByVal pTestCd As String) As String
    Dim objFTestCd  As clsIISTestCdFull '�˻��ڵ� Ŭ����
    Dim strLastDt   As String           '�˻��ڵ��� �ֱ� ������
    
    If mFTestCds.Count = 0 Then Exit Function
    For Each objFTestCd In mFTestCds
        If strLastDt < objFTestCd.Applydt Then
            strLastDt = objFTestCd.Applydt
        End If
    Next
    Set objFTestCd = Nothing
    GetTestCdLastApplyDt = strLastDt
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ��ü�ڵ��� �����ֱ� �������� ��ȸ
'   �μ� :
'       1.mSpcCd  : ��ü�ڵ�
'   ��ȯ :
'       1.�������� �����ϸ� ���� �ֱ� �������� ��ȯ
'       2.�������� ������ Null�� ��ȯ
'-----------------------------------------------------------------------------'
Public Function GetSpcCdLastApplyDt(ByVal pSpcCd As String)
    Dim objTSpc     As clsIISTSpc   '������ü Ŭ����
    Dim strLastDt   As String       '��ü�ڵ��� �ֱ� ������
    
    If mTSpcs.Count = 0 Then Exit Function
    For Each objTSpc In mTSpcs
        If pSpcCd = objTSpc.SpcCd And strLastDt < objTSpc.Applydt Then
            strLastDt = objTSpc.Applydt
        End If
    Next
    GetSpcCdLastApplyDt = strLastDt
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ����ġ�� �����ֱ� ������ ��ȸ
'   �μ� :
'       1.mSpcCd  : ��ü�ڵ�
'   ��ȯ :
'       1.�������� �����ϸ� ���� �ֱ� �������� ��ȯ
'       2.�������� ������ Null�� ��ȯ
'-----------------------------------------------------------------------------'
Public Function GetRefLastApplyDt(ByVal pSpcCd As String)
    Dim objRef      As clsIISRef    '����ġ Ŭ����
    Dim strLastDt   As String       '����ġ�� �ֱ� ������
    
    If mRefs.Count = 0 Then Exit Function
    For Each objRef In mRefs
        If pSpcCd = objRef.SpcCd And strLastDt < objRef.Applydt Then
            strLastDt = objRef.Applydt
        End If
    Next
    GetRefLastApplyDt = strLastDt
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻��ڵ� ����Ʈ ��ȸ
'-----------------------------------------------------------------------------'
Public Function GetTestCdList()
    Dim objPro      As clsIISProgress
    Dim Rs          As ADODB.Recordset
    Dim strTestCd   As String           '�˻��ڵ�
    Dim strTestNm   As String           '�˻��(5�ڸ�)
    Dim i           As Long

On Error GoTo Errors
    mTestCds.RemoveAll
    Set Rs = New ADODB.Recordset
    Rs.Open mSQL.SelectTestCd, DbCon, adOpenDynamic, adLockReadOnly, adCmdText
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    i = 1
    Set objPro = New clsIISProgress
    With objPro
        .StatusBar = StatusBar
        .PanelPos = 2
        .Max = Rs.RecordCount
    End With
    
    Do Until Rs.EOF
        strTestCd = Rs.Fields("TESTCD").Value
        strTestNm = Rs.Fields("TESTNM5").Value & ""
        
        If mTestCds.Exist(strTestCd) = False Then
            mTestCds.Add strTestCd, i, strTestNm
            i = i + 1
        End If
        
        objPro.Value = i
        Rs.MoveNext
    Loop
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Set objPro = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Set objPro = Nothing
    Error.SetLog App.EXEName, "clsIISTMaster", "GetTestCdList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� �˻��׸� ������ȸ
'   �μ� :
'       1.pTestCd : �˻��ڵ�
'   ��ȯ : �˻��ڵ�(��ü) �÷���
'-----------------------------------------------------------------------------'
Public Function GetTestCdInfo(ByVal pTestCd As String) As clsIISTestCdFulls
    Dim Rs          As ADODB.Recordset
    Dim strApplyDt  As String       '������
    Dim strExpireDt As String       '�����
    Dim strTestNm5  As String       '�˻��(5�ڸ�)
    Dim strTestNm10 As String       '�˻��(10�ڸ�)
    Dim strTestNm   As String       '�˻��(��ü)
    Dim strWorkarea As String       'Workarea�ڵ�
    Dim strRstType  As String       '�������(Null: �Ϲ�, F: Free, N: Numeric, A: Alpha)
    Dim strRstDiv   As String       'Alternative, Require ����(A: Alternative, R: Require)
    Dim strPanelFg  As String       'ó�汸��(Null: �����׸�, G: �׷��׸�, D: ���׸�)
    Dim strDetailFg As String       'Detail �׸񿩺�(Null: ����, *:�����׸� ���ڵ�)
    Dim lngRptSeq   As Long         '��¼���
    
On Error GoTo Errors
    mFTestCds.RemoveAll
    Set Rs = DbCon.Execute(mSQL.SelectTestCdX(pTestCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    With mFTestCds
        Do Until Rs.EOF
            strApplyDt = Rs.Fields("APPLYDT").Value
            strExpireDt = Rs.Fields("EXPIREDT").Value & ""
            strTestNm5 = Rs.Fields("TESTNM5").Value
            strTestNm10 = Rs.Fields("TESTNM10").Value & ""
            strTestNm = Rs.Fields("TESTNM").Value
            strWorkarea = Rs.Fields("WORKAREA").Value
            strRstType = Rs.Fields("RSTTYPE").Value & ""
            strRstDiv = Rs.Fields("RSTDIV").Value & ""
            strPanelFg = Rs.Fields("PANELFG").Value & ""
            strDetailFg = Rs.Fields("DETAILFG").Value & ""
            lngRptSeq = CLng(Rs.Fields("RPTSEQ").Value)
        
            .Add pTestCd, strApplyDt, strWorkarea, strTestNm, strExpireDt, strTestNm5, strTestNm10, _
                strRstType, strRstDiv, strPanelFg, strDetailFg, lngRptSeq
            Rs.MoveNext
        Loop
    End With
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Set GetTestCdInfo = mFTestCds
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISTMaster", "GetTestCdInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� �˻��׸��� ������ü ���� ��ȸ
'   �μ� :
'       1.pTestCd : �˻��ڵ�
'   ��ȯ : ������ü �÷���
'-----------------------------------------------------------------------------'
Public Function GetSpcInfo(ByVal pTestCd As String) As clsIISTSpcs
    Dim Rs              As ADODB.Recordset
    Dim strSpcCd        As String       '��ü�ڵ�
    Dim strSpcNm        As String       '��ü��
    Dim strApplyDt      As String       '������
    Dim strExpireDt     As String       '�����
    Dim strUnit         As String       '����
    Dim lngAvalVal      As Long         '��ȿ����
    Dim strPanicFg      As String       'Panic Check(0:No, 1:Yes)
    Dim sngPanicFrVal   As Single       'Panic From Value
    Dim sngPanicToVal   As Single       'Panic To Value
    Dim strDeltaFg      As String       'Delta Check(0:No, 1:Yes)
    Dim lngDeltaFrVal   As Long         'Delta From Value
    Dim lngDeltaToVal   As Long         'Delta To Value
    
On Error GoTo Errors
    mTSpcs.RemoveAll
    Set Rs = DbCon.Execute(mSQL.SelectTSpcCd(pTestCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    With mTSpcs
        Do Until Rs.EOF
            strSpcCd = Rs.Fields("SPCCD").Value
            strSpcNm = Rs.Fields("SPCNM").Value
            strApplyDt = Rs.Fields("APPLYDT").Value
            strExpireDt = Rs.Fields("EXPIREDT").Value & ""
            strUnit = Rs.Fields("UNIT").Value & ""
            lngAvalVal = CLng(Rs.Fields("AVALVAL").Value)
            strPanicFg = Rs.Fields("PANICFG").Value & ""
            sngPanicFrVal = CSng(Rs.Fields("PANICFRVAL").Value)
            sngPanicToVal = CSng(Rs.Fields("PANICTOVAL").Value)
            strDeltaFg = Rs.Fields("DELTAFG").Value & ""
            lngDeltaFrVal = CLng(Rs.Fields("DELTAFRVAL").Value)
            lngDeltaToVal = CLng(Rs.Fields("DELTATOVAL").Value)
            
            .Add pTestCd, strSpcCd, strSpcNm, strApplyDt, strExpireDt, strUnit, lngAvalVal, _
                strPanicFg, sngPanicFrVal, sngPanicToVal, strDeltaFg, lngDeltaFrVal, lngDeltaToVal
            Rs.MoveNext
        Loop
    End With
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Set GetSpcInfo = mTSpcs
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISTMaster", "GetSpcInfo", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� �˻��׸�, ��ü�ڵ��� ����ġ ������ ��ȸ
'   �μ� :
'       1.pTestCd : �˻��ڵ�
'       2.pSpcCd  : ��ü�ڵ�
'-----------------------------------------------------------------------------'
Public Function GetRefList(ByVal pTestCd As String, ByVal pSpcCd As String) As clsIISRefs
    Dim Rs          As ADODB.Recordset
    Dim strSex      As String       '����
    Dim lngAgeFr    As Long         'From Age
    Dim lngAgeTo    As Long         'To Age
    Dim strApplyDt  As String       '������
    Dim strExpireDt As String       '�����
    Dim sngRefFrVal As Single       'From Reference Value
    Dim sngRefToVal As Single       'To Reference Value
    Dim strRefCd    As String       'Alpha��� ����ġ
    
On Error GoTo Errors
    mRefs.RemoveAll
    Set Rs = DbCon.Execute(mSQL.SelectReference(pTestCd, pSpcCd), , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    With mRefs
        Do Until Rs.EOF
            strSex = Rs.Fields("SEX").Value
            lngAgeFr = CLng(Rs.Fields("AGEFR").Value)
            lngAgeTo = CLng(Rs.Fields("AGETO").Value)
            strApplyDt = Rs.Fields("APPLYDT").Value
            strExpireDt = Rs.Fields("EXPIREDT").Value & ""
            sngRefFrVal = CSng(Rs.Fields("REFFRVAL").Value)
            sngRefToVal = CSng(Rs.Fields("REFTOVAL").Value)
            strRefCd = Rs.Fields("REFCD").Value & ""
            
            .Add pTestCd, pSpcCd, strSex, lngAgeFr, lngAgeTo, strApplyDt, strExpireDt, _
                sngRefFrVal, sngRefToVal, strRefCd
            Rs.MoveNext
        Loop
    End With
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Set GetRefList = mRefs
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISTMaster", "GetRefList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻��ڵ� �߰�
'   �μ� :
'       1.pTestCd    : �˻��ڵ�
'       2.pApplyDt   : ������
'       3.pExpireDt  : �����
'       4.pTestNm5   : �˻��(5�ڸ�)
'       5.pTestNm10  : �˻��(10�ڸ�)
'       6.pTestNm    : �˻��(��ü)
'       7.pWorkarea  : Workarea �ڵ�
'       8.pRstType   : �������(Null: �Ϲ�, F: Free, N: Numeric, A: Alpha)
'       9.pRstDiv    : Alternative, Require ����(A: Alternative, R: Require)
'       10.pPanelFg  : ó�汸��(Null: �����׸�, G: �׷��׸�, D: ���׸�)
'       11.pDetailFg : Detail �׸񿩺�(Null: ����, *:�����׸� ���ڵ�)
'       12.pRptSeq   : ��¼���
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddTestCd(ByVal pTestCd As String, ByVal pApplyDt As String, _
                          ByVal pExpireDt As String, ByVal pTestNm5 As String, _
                          ByVal pTestNm10 As String, ByVal pTestNm As String, _
                          ByVal pWorkarea As String, ByVal pRstType As String, _
                          ByVal pRstDiv As String, ByVal pPanelFg As String, _
                          ByVal pDetailFg As String, ByVal pRptSeq As Long) As Boolean
                          
    Dim objFTestCd As clsIISTestCdFull
    
    Set objFTestCd = New clsIISTestCdFull
    With objFTestCd
        .TestCd = pTestCd
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .TestNm5 = pTestNm5
        .TestNm10 = pTestNm10
        .TestNm = pTestNm
        .Workarea = pWorkarea
        .RstDiv = pRstDiv
        .PanelFg = pPanelFg
        .DetailFg = pDetailFg
        .RptSeq = pRptSeq
        If .AddTestCd Then
            AddTestCd = True
        End If
    End With
    Set objFTestCd = Nothing
    
    '## �ű� �˻��ڵ��� ��� mTestCds �÷��ǿ� �˻��ڵ� ���� �߰��� �÷��� ����
    '## NOTE:���� �˰������� ����߾�� �ϳ� ���� Db���� �ٽ� �ε�
    If mTestCds.Exist(pTestCd) = False Then
        Dim objSort As clsIISSort

        mTestCds.Add pTestCd, mTestCds.Count + 1, pTestNm5
        Set objSort = New clsIISSort
        Call objSort.Sort(mTestCds)
        
        Set objSort = Nothing
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻��ڵ� ����(�˻��ڵ� ������ ���� �˻��ڵ�+�������� �����ϰ� �ٽ��Է�)
'   �μ� :
'       1.pTestCd    : �˻��ڵ�
'       2.pApplyDt   : ������
'       3.pExpireDt  : �����
'       4.pTestNm5   : �˻��(5�ڸ�)
'       5.pTestNm10  : �˻��(10�ڸ�)
'       6.pTestNm    : �˻��(��ü)
'       7.pWorkarea  : Workarea �ڵ�
'       8.pRstType   : �������(Null: �Ϲ�, F: Free, N: Numeric, A: Alpha)
'       9.pRstDiv    : Alternative, Require ����(A: Alternative, R: Require)
'       10.pPanelFg  : ó�汸��(Null: �����׸�, G: �׷��׸�, D: ���׸�)
'       11.pDetailFg : Detail �׸񿩺�(Null: ����, *:�����׸� ���ڵ�)
'       12.pRptSeq   : ��¼���
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function ModifyTestCd(ByVal pTestCd As String, ByVal pApplyDt As String, _
                             ByVal pExpireDt As String, ByVal pTestNm5 As String, _
                             ByVal pTestNm10 As String, ByVal pTestNm As String, _
                             ByVal pWorkarea As String, ByVal pRstType As String, _
                             ByVal pRstDiv As String, ByVal pPanelFg As String, _
                             ByVal pDetailFg As String, ByVal pRptSeq As Long) As Boolean

    Dim objFTestCd As clsIISTestCdFull
    
    Set objFTestCd = New clsIISTestCdFull
    With objFTestCd
        .TestCd = pTestCd
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .TestNm5 = pTestNm5
        .TestNm10 = pTestNm10
        .TestNm = pTestNm
        .Workarea = pWorkarea
        .RstType = pRstType
        .RstDiv = pRstDiv
        .PanelFg = pPanelFg
        .DetailFg = pDetailFg
        .RptSeq = pRptSeq
        If .ModifyTestCd Then
            ModifyTestCd = True
        End If
    End With
    Set objFTestCd = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �˻��ڵ� ����
'   �μ� :
'       1.pTestCd    : �˻��ڵ�
'       2.pApplyDt   : ������
'   ��ȯ :
'       1. : ������ �����ϸ� '-1' ��ȯ
'       2. : ���������� �����ϸ� ���� ������ ������ ��ȯ
'-----------------------------------------------------------------------------'
Public Function RemoveTestCd(ByVal pTestCd As String, ByVal pApplyDt As String) As Long
    Dim objFTestCd As clsIISTestCdFull
    
    Set objFTestCd = New clsIISTestCdFull
    With objFTestCd
        .TestCd = pTestCd
        .Applydt = pApplyDt
        If .DelTestCd = False Then
            RemoveTestCd = -1
            Exit Function
        End If
    End With
    Set objFTestCd = Nothing
    
    '## ���� ������ ������ 0�̸� mTestCds �÷��ǿ��� ����
    '## ���� ������ ������ ��� �ľ�?
    '   - mFTestCds �÷��ǿ��� ���� �˻��ڵ�+������ ������ Count=0 �϶�
    '     (� �˻��ڵ��� �������� �����Ϸ��� mFTestCds �÷��ǿ� �׻� �ش� �˻��ڵ�+������
    '      ����Ʈ�� ��� �־�� �Ѵ�.)
    Call mFTestCds.Remove(pTestCd, pApplyDt)
    If mFTestCds.Count = 0 Then
        Call mTestCds.Remove(pTestCd)
        Call mTestCds.SortAdd(mTestCds)
        RemoveTestCd = 0
    Else
        RemoveTestCd = mFTestCds.Count
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������ü �߰�
'   �μ� :
'       1.pTestCd       : �˻��ڵ�
'       2.pSpcCd        : ��ü�ڵ�
'       3.pApplyDt      : ������
'       4.pExpireDt     : �����
'       5.pUnit         : �������
'       6.pAvalVal      : ��ȿ����
'       7.pPanicFg      : Panic Check (0:No, 1:Yes)
'       8.pPanicFrVal   : Panic From Value
'       9.pPanicToVal   : Panic To Value
'       10.pDeltaFg     : Delta Check (0:No, 1:Yes)
'       11.pDeltaFrVal  : Panic From Value
'       12.pDeltaToVal  : Panic To Value
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddSpcCd(ByVal pTestCd As String, ByVal pSpcCd As String, _
                         ByVal pApplyDt As String, ByVal pExpireDt As String, _
                         ByVal pUnit As String, ByVal pAvalVal As Long, _
                         ByVal pPanicFg As String, ByVal pPanicFrVal As Single, _
                         ByVal pPanicToVal As Single, ByVal pDeltaFg As String, _
                         ByVal pDeltaFrVal As Long, ByVal pDeltaToVal As Long) As Boolean

    Dim objTSpc As clsIISTSpc
    
    Set objTSpc = New clsIISTSpc
    With objTSpc
        .TestCd = pTestCd
        .SpcCd = pSpcCd
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .Unit = pUnit
        .AvalVal = pAvalVal
        .PanicFg = pPanicFg
        .PanicFrVal = pPanicFrVal
        .PanicToVal = pPanicToVal
        .DeltaFg = pDeltaFg
        .DeltaFrVal = pDeltaFrVal
        .DeltaToVal = pDeltaToVal
        If .AddSpcCd Then AddSpcCd = True
    End With
    Set objTSpc = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������ü ����
'   �μ� :
'       1.pTestCd       : �˻��ڵ�
'       2.pSpcCd        : ��ü�ڵ�
'       3.pApplyDt      : ������
'       4.pExpireDt     : �����
'       5.pUnit         : �������
'       6.pAvalVal      : ��ȿ����
'       7.pPanicFg      : Panic Check (0:No, 1:Yes)
'       8.pPanicFrVal   : Panic From Value
'       9.pPanicToVal   : Panic To Value
'       10.pDeltaFg     : Delta Check (0:No, 1:Yes)
'       11.pDeltaFrVal  : Panic From Value
'       12.pDeltaToVal  : Panic To Value
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function ModifySpcCd(ByVal pTestCd As String, ByVal pSpcCd As String, _
                            ByVal pApplyDt As String, ByVal pExpireDt As String, _
                            ByVal pUnit As String, ByVal pAvalVal As Long, _
                            ByVal pPanicFg As String, ByVal pPanicFrVal As Single, _
                            ByVal pPanicToVal As Single, ByVal pDeltaFg As String, _
                            ByVal pDeltaFrVal As Long, ByVal pDeltaToVal As Long) As Boolean
                            
    Dim objTSpc As clsIISTSpc
    
    Set objTSpc = New clsIISTSpc
    With objTSpc
        .TestCd = pTestCd
        .SpcCd = pSpcCd
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .Unit = pUnit
        .AvalVal = pAvalVal
        .PanicFg = pPanicFg
        .PanicFrVal = pPanicFrVal
        .PanicToVal = pPanicToVal
        .DeltaFg = pDeltaFg
        .DeltaFrVal = pDeltaFrVal
        .DeltaToVal = pDeltaToVal
        If .ModifySpcCd Then ModifySpcCd = True
    End With
    Set objTSpc = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������ü ����
'   �μ� :
'       1.pTestCd  : �˻��ڵ�
'       2.pSpcCd   : ��ü�ڵ�
'       3.pApplyDt : ������
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function RemoveSpcCd(ByVal pTestCd As String, ByVal pSpcCd As String, _
                            ByVal pApplyDt As String) As Boolean
                            
    Dim objTSpc As clsIISTSpc
    
    Set objTSpc = New clsIISTSpc
    With objTSpc
        .TestCd = pTestCd
        .SpcCd = pSpcCd
        .Applydt = pApplyDt
        If .DelSpcCd Then RemoveSpcCd = True
    End With
    Set objTSpc = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : ����ġ �߰�
'   �μ� :
'       1.pTestCd   : �˻��ڵ�
'       2.pSpcCd    : ��ü�ڵ�
'       3.pSex      : ����
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : ������
'       7.pExpireDt : �����
'       8.pRefFrVal : From Reference Value
'       9.pRefToVal : To Reference Value
'       10.pRefCd   : Alpha��� ����ġ
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function AddRef(ByVal pTestCd As String, ByVal pSpcCd As String, _
                       ByVal pSex As String, ByVal pAgeFr As Long, _
                       ByVal pAgeTo As Long, ByVal pApplyDt As String, _
                       ByVal pExpireDt As String, ByVal pRefFrVal As Single, _
                       ByVal pRefToVal As Single, ByVal pRefCd As String) As Boolean

    Dim objRef As clsIISRef

    Set objRef = New clsIISRef
    With objRef
        .TestCd = pTestCd
        .SpcCd = pSpcCd
        .Sex = pSex
        .AgeFr = pAgeFr
        .AgeTo = pAgeTo
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .RefFrVal = pRefFrVal
        .RefToVal = pRefToVal
        .Refcd = pRefCd
        If .AddRef Then AddRef = True
    End With
    Set objRef = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : ����ġ ����
'   �μ� :
'       1.pTestCd   : �˻��ڵ�
'       2.pSpcCd    : ��ü�ڵ�
'       3.pSex      : ����
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : ������
'       7.pExpireDt : �����
'       8.pRefFrVal : From Reference Value
'       9.pRefToVal : To Reference Value
'       10. pRefCd  : Alpha��� ����ġ
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function ModifyRef(ByVal pTestCd As String, ByVal pSpcCd As String, _
                          ByVal pSex As String, ByVal pAgeFr As Long, _
                          ByVal pAgeTo As Long, ByVal pApplyDt As String, _
                          ByVal pExpireDt As String, ByVal pRefFrVal As Single, _
                          ByVal pRefToVal As Single, ByVal pRefCd As String) As Boolean

    Dim objRef As clsIISRef

    Set objRef = New clsIISRef
    With objRef
        .TestCd = pTestCd
        .SpcCd = pSpcCd
        .Sex = pSex
        .AgeFr = pAgeFr
        .AgeTo = pAgeTo
        .Applydt = pApplyDt
        .ExpireDt = pExpireDt
        .RefFrVal = pRefFrVal
        .RefToVal = pRefToVal
        .Refcd = pRefCd
        If .ModifyRef Then ModifyRef = True
    End With
    Set objRef = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : ����ġ ����
'   �μ� :
'       1.pTestCd   : �˻��ڵ�
'       2.pSpcCd    : ��ü�ڵ�
'       3.pSex      : ����
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : ������
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function RemoveRef(ByVal pTestCd As String, ByVal pSpcCd As String, _
                          ByVal pSex As String, ByVal pAgeFr As Long, _
                          ByVal pAgeTo As Long, ByVal pApplyDt As String) As Boolean

    Dim objRef As clsIISRef

    Set objRef = New clsIISRef
    With objRef
        .TestCd = pTestCd
        .SpcCd = pSpcCd
        .Sex = pSex
        .AgeFr = pAgeFr
        .AgeTo = pAgeTo
        .Applydt = pApplyDt
        If .DelRef Then RemoveRef = True
    End With
    Set objRef = Nothing
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
    Set mTestCds = New clsIISTestCds
    Set mFTestCds = New clsIISTestCdFulls
    Set mTSpcs = New clsIISTSpcs
    Set mRefs = New clsIISRefs
End Sub

Private Sub Class_Terminate()
    Set mRefs = Nothing
    Set mTSpcs = Nothing
    Set mFTestCds = Nothing
    Set mTestCds = Nothing
    Set mSQL = New clsIISMasterSql
End Sub