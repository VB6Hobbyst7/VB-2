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
'   파일명  : clsIISTMaster.cls

'   내  용  : 검사코드 마스터 클래스
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL        As clsIISMasterSql

Private mTestCds    As clsIISTestCds        '검사코드 컬렉션
Private mFTestCds   As clsIISTestCdFulls    '검사코드(전체) 컬렉션
Private mTSpcs      As clsIISTSpcs          '지정검체 컬렉션
Private mRefs       As clsIISRefs           '참고치 컬렉션

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
'   기능 : 해당 검사코드의 다음 검사코드를 조회
'          다음 검사코드가 없으면 현재 검사코드 리턴
'   인수 :
'       1.pTestCd : 검사코드
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
'   기능 : 해당 검사코드의 이전 검사코드를 조회
'          이전 검사코드가 없으면 현재 검사코드 리턴
'   인수 :
'       1.pTestCd : 검사코드
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
'   기능 : 해당 검사코드의 존재유무를 조회
'   인수 :
'       1.pTestCd : 검사코드
'   반환 : True(존재), False(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pTestCd As String) As Boolean
    If mTestCds.Exist(pTestCd) Then
        Exist = True
    Else
        Exist = False
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검체코드의 존재유무를 조회
'   인수 :
'       1.pSpcCd  : 검체코드
'   반환 : True(존재), False(비존재)
'-----------------------------------------------------------------------------'
Public Function ExistSpcCd(ByVal pSpcCd As String) As Boolean
    Dim objTSpc As clsIISTSpc   '지정검체 클래스
    
    If mTSpcs.Count = 0 Then Exit Function
    For Each objTSpc In mTSpcs
        If pSpcCd = objTSpc.SpcCd Then
            ExistSpcCd = True
            Exit For
        End If
    Next
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 참고치의 존재유무를 조회
'   인수 :
'       1.pSpcCd  : 검체코드
'   반환 : True(존재), False(비존재)
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
'   기능 : 해당 검사코드의 검사명을 조회
'   인수 :
'       1.pTestCd : 검사코드
'   반환 : 검사명
'-----------------------------------------------------------------------------'
Public Function GetTestNm(ByVal pTestNm As String) As String
    '## 검사코드가 반드시 존재해야 한다.
    GetTestNm = mTestCds(pTestNm).TestNm
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사코드의 가장최근 적용일을 조회
'   인수 :
'       1.mTestCd : 검사코드
'   반환 :
'       1.적용일이 존재하면 가장 최근 적용일을 반환
'       2.적용일이 없으면 Null을 반환
'-----------------------------------------------------------------------------'
Public Function GetTestCdLastApplyDt(ByVal pTestCd As String) As String
    Dim objFTestCd  As clsIISTestCdFull '검사코드 클래스
    Dim strLastDt   As String           '검사코드의 최근 적용일
    
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
'   기능 : 해당 검체코드의 가장최근 적용일을 조회
'   인수 :
'       1.mSpcCd  : 검체코드
'   반환 :
'       1.적용일이 존재하면 가장 최근 적용일을 반환
'       2.적용일이 없으면 Null을 반환
'-----------------------------------------------------------------------------'
Public Function GetSpcCdLastApplyDt(ByVal pSpcCd As String)
    Dim objTSpc     As clsIISTSpc   '지정검체 클래스
    Dim strLastDt   As String       '검체코드의 최근 적용일
    
    If mTSpcs.Count = 0 Then Exit Function
    For Each objTSpc In mTSpcs
        If pSpcCd = objTSpc.SpcCd And strLastDt < objTSpc.Applydt Then
            strLastDt = objTSpc.Applydt
        End If
    Next
    GetSpcCdLastApplyDt = strLastDt
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 참고치중 가장최근 적용일 조회
'   인수 :
'       1.mSpcCd  : 검체코드
'   반환 :
'       1.적용일이 존재하면 가장 최근 적용일을 반환
'       2.적용일이 없으면 Null을 반환
'-----------------------------------------------------------------------------'
Public Function GetRefLastApplyDt(ByVal pSpcCd As String)
    Dim objRef      As clsIISRef    '참고치 클래스
    Dim strLastDt   As String       '참고치의 최근 적용일
    
    If mRefs.Count = 0 Then Exit Function
    For Each objRef In mRefs
        If pSpcCd = objRef.SpcCd And strLastDt < objRef.Applydt Then
            strLastDt = objRef.Applydt
        End If
    Next
    GetRefLastApplyDt = strLastDt
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사코드 리스트 조회
'-----------------------------------------------------------------------------'
Public Function GetTestCdList()
    Dim objPro      As clsIISProgress
    Dim Rs          As ADODB.Recordset
    Dim strTestCd   As String           '검사코드
    Dim strTestNm   As String           '검사명(5자리)
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
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사항목 정보조회
'   인수 :
'       1.pTestCd : 검사코드
'   반환 : 검사코드(전체) 컬렉션
'-----------------------------------------------------------------------------'
Public Function GetTestCdInfo(ByVal pTestCd As String) As clsIISTestCdFulls
    Dim Rs          As ADODB.Recordset
    Dim strApplyDt  As String       '적용일
    Dim strExpireDt As String       '폐기일
    Dim strTestNm5  As String       '검사명(5자리)
    Dim strTestNm10 As String       '검사명(10자리)
    Dim strTestNm   As String       '검사명(전체)
    Dim strWorkarea As String       'Workarea코드
    Dim strRstType  As String       '결과유형(Null: 일반, F: Free, N: Numeric, A: Alpha)
    Dim strRstDiv   As String       'Alternative, Require 유형(A: Alternative, R: Require)
    Dim strPanelFg  As String       '처방구분(Null: 단일항목, G: 그룹항목, D: 상세항목)
    Dim strDetailFg As String       'Detail 항목여부(Null: 없음, *:단일항목 자코드)
    Dim lngRptSeq   As Long         '출력순서
    
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
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사항목의 지정검체 정보 조회
'   인수 :
'       1.pTestCd : 검사코드
'   반환 : 지정검체 컬렉션
'-----------------------------------------------------------------------------'
Public Function GetSpcInfo(ByVal pTestCd As String) As clsIISTSpcs
    Dim Rs              As ADODB.Recordset
    Dim strSpcCd        As String       '검체코드
    Dim strSpcNm        As String       '검체명
    Dim strApplyDt      As String       '적용일
    Dim strExpireDt     As String       '폐기일
    Dim strUnit         As String       '단위
    Dim lngAvalVal      As Long         '유효숫자
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
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사항목, 검체코드의 참고치 정보를 조회
'   인수 :
'       1.pTestCd : 검사코드
'       2.pSpcCd  : 검체코드
'-----------------------------------------------------------------------------'
Public Function GetRefList(ByVal pTestCd As String, ByVal pSpcCd As String) As clsIISRefs
    Dim Rs          As ADODB.Recordset
    Dim strSex      As String       '성별
    Dim lngAgeFr    As Long         'From Age
    Dim lngAgeTo    As Long         'To Age
    Dim strApplyDt  As String       '적용일
    Dim strExpireDt As String       '폐기일
    Dim sngRefFrVal As Single       'From Reference Value
    Dim sngRefToVal As Single       'To Reference Value
    Dim strRefCd    As String       'Alpha결과 참고치
    
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
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사코드 추가
'   인수 :
'       1.pTestCd    : 검사코드
'       2.pApplyDt   : 적용일
'       3.pExpireDt  : 폐기일
'       4.pTestNm5   : 검사명(5자리)
'       5.pTestNm10  : 검사명(10자리)
'       6.pTestNm    : 검사명(전체)
'       7.pWorkarea  : Workarea 코드
'       8.pRstType   : 결과유형(Null: 일반, F: Free, N: Numeric, A: Alpha)
'       9.pRstDiv    : Alternative, Require 유형(A: Alternative, R: Require)
'       10.pPanelFg  : 처방구분(Null: 단일항목, G: 그룹항목, D: 상세항목)
'       11.pDetailFg : Detail 항목여부(Null: 없음, *:단일항목 자코드)
'       12.pRptSeq   : 출력순서
'   반환 : True(성공), False(실패)
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
    
    '## 신규 검사코드인 경우 mTestCds 컬렉션에 검사코드 정보 추가후 컬렉션 정렬
    '## NOTE:정렬 알고리즘을 사용했어야 하나 몰라서 Db에서 다시 로드
    If mTestCds.Exist(pTestCd) = False Then
        Dim objSort As clsIISSort

        mTestCds.Add pTestCd, mTestCds.Count + 1, pTestNm5
        Set objSort = New clsIISSort
        Call objSort.Sort(mTestCds)
        
        Set objSort = Nothing
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 검사코드 수정(검사코드 수정시 기존 검사코드+적용일을 삭제하고 다시입력)
'   인수 :
'       1.pTestCd    : 검사코드
'       2.pApplyDt   : 적용일
'       3.pExpireDt  : 폐기일
'       4.pTestNm5   : 검사명(5자리)
'       5.pTestNm10  : 검사명(10자리)
'       6.pTestNm    : 검사명(전체)
'       7.pWorkarea  : Workarea 코드
'       8.pRstType   : 결과유형(Null: 일반, F: Free, N: Numeric, A: Alpha)
'       9.pRstDiv    : Alternative, Require 유형(A: Alternative, R: Require)
'       10.pPanelFg  : 처방구분(Null: 단일항목, G: 그룹항목, D: 상세항목)
'       11.pDetailFg : Detail 항목여부(Null: 없음, *:단일항목 자코드)
'       12.pRptSeq   : 출력순서
'   반환 : True(성공), False(실패)
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
'   기능 : 검사코드 삭제
'   인수 :
'       1.pTestCd    : 검사코드
'       2.pApplyDt   : 적용일
'   반환 :
'       1. : 삭제중 실패하면 '-1' 반환
'       2. : 성공적으로 삭제하면 남은 적용일 개수를 반환
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
    
    '## 남은 적용일 개수가 0이면 mTestCds 컬렉션에서 삭제
    '## 남은 적용일 개수를 어떻게 파악?
    '   - mFTestCds 컬렉션에서 현재 검사코드+적용일 삭제후 Count=0 일때
    '     (어떤 검사코드의 적용일을 삭제하려면 mFTestCds 컬렉션에 항상 해당 검사코드+적용일
    '      리스트가 담겨 있어야 한다.)
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
'   기능 : 지정검체 추가
'   인수 :
'       1.pTestCd       : 검사코드
'       2.pSpcCd        : 검체코드
'       3.pApplyDt      : 적용일
'       4.pExpireDt     : 폐기일
'       5.pUnit         : 결과단위
'       6.pAvalVal      : 유효숫자
'       7.pPanicFg      : Panic Check (0:No, 1:Yes)
'       8.pPanicFrVal   : Panic From Value
'       9.pPanicToVal   : Panic To Value
'       10.pDeltaFg     : Delta Check (0:No, 1:Yes)
'       11.pDeltaFrVal  : Panic From Value
'       12.pDeltaToVal  : Panic To Value
'   반환 : True(성공), False(실패)
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
'   기능 : 지정검체 수정
'   인수 :
'       1.pTestCd       : 검사코드
'       2.pSpcCd        : 검체코드
'       3.pApplyDt      : 적용일
'       4.pExpireDt     : 폐기일
'       5.pUnit         : 결과단위
'       6.pAvalVal      : 유효숫자
'       7.pPanicFg      : Panic Check (0:No, 1:Yes)
'       8.pPanicFrVal   : Panic From Value
'       9.pPanicToVal   : Panic To Value
'       10.pDeltaFg     : Delta Check (0:No, 1:Yes)
'       11.pDeltaFrVal  : Panic From Value
'       12.pDeltaToVal  : Panic To Value
'   반환 : True(성공), False(실패)
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
'   기능 : 지정검체 삭제
'   인수 :
'       1.pTestCd  : 검사코드
'       2.pSpcCd   : 검체코드
'       3.pApplyDt : 적용일
'   반환 : True(성공), False(실패)
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
'   기능 : 참고치 추가
'   인수 :
'       1.pTestCd   : 검사코드
'       2.pSpcCd    : 검체코드
'       3.pSex      : 성별
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : 적용일
'       7.pExpireDt : 폐기일
'       8.pRefFrVal : From Reference Value
'       9.pRefToVal : To Reference Value
'       10.pRefCd   : Alpha결과 참고치
'   반환 : True(성공), False(실패)
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
'   기능 : 참고치 수정
'   인수 :
'       1.pTestCd   : 검사코드
'       2.pSpcCd    : 검체코드
'       3.pSex      : 성별
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : 적용일
'       7.pExpireDt : 폐기일
'       8.pRefFrVal : From Reference Value
'       9.pRefToVal : To Reference Value
'       10. pRefCd  : Alpha결과 참고치
'   반환 : True(성공), False(실패)
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
'   기능 : 참고치 삭제
'   인수 :
'       1.pTestCd   : 검사코드
'       2.pSpcCd    : 검체코드
'       3.pSex      : 성별
'       4.pAgeFr    : From Age
'       5.pAgeTo    : To Age
'       6.pApplyDt  : 적용일
'   반환 : True(성공), False(실패)
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
