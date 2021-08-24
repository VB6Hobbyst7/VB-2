VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   파일명  : clsIISInterface.cls (자사 LIS와 조인할때)
'   작성자  : 이상대
'   내  용  : 인터페이스 클래스
'   작성일  : 2004-03-31
'   버  전  :
'       1. 1.1.5: 이상대(2005-04-12)
'       2. 1.1.10: 이상대(2005-07-06)
'       3. 1.1.12: 이상대(2005-07-28)
'       4. 1.1.14: 이상대(2005-10-11)
'   메  모  :
'       1. IIS-메모.doc의 2004-04-09의 1번 참고
'-----------------------------------------------------------------------------'

Option Explicit

'## Log Enum
Public Enum LogEnum
    ccPCLog = 0         'PC에서 송신한 Log
    ccEqpLog = 1        '장비에서 송신한 Log
End Enum

Private mIntNms     As clsIISIntNms         '장비별 검사항목 컬렉션 클래스(LIS검사항목 기준)
Private mAccInfos   As clsIISAccInfos       '접수내역 컬렉션 클래스
Private mBuffers    As clsIISBuffers        '버퍼 컬렉션 클래스

'2007.05.18 osw 수정 - 하나의 I/F 프로그램에서 다중의 장비와 연결되었을때 장비코드를 조정이 가능하게
'                      I/F프로그램에서 핸들링이 가능하도록 Public 형태로 수정
'Private mEqpCd      As String           '장비코드
'Private mEqpKey     As String           '장비키
Public mEqpCd       As String           '장비코드
Public mEqpKey      As String           '장비키
Private mEqpType    As EqpCommTypeEnum  '장비통신유형
Private mBarcode    As BarcodeUseEnum   '바코드 사용유무
Private mAutoVfy    As AutoVfyEnum      '자동결과등록 사용유무
Private mHLCheck    As HLCheckEnum      '자동결과등록시 H/L에 처리여부
Private mBarPos     As BarPosEnum       '바코드 리더기 위치
Private mStoreDt    As Long             '보관일
Private mFileNum    As Integer          '로그파일의 File Number

Private mPhase      As Long             '통신상태(전체)
Private mSndPhase   As Long             '통신상태(송신)
Private mBufCnt     As Long             '버퍼순서
Private mState      As String           '현재 통신상태
Private mFrameNo    As Long             'ASTM 방식의 FrameNo
Private mIsETB      As Boolean          'ASTM 방식의 ETB 수신여부
Private mACKFg      As Long             'ACK,ETX 전송여부(0:ETX전송후,1:ACK전송후) -- CX 검사장비에서 사용

'## 1.1.10: 이상대(2005-07-06)
'   - 수신상태 변수추가
Private mRcvPhase   As Long             '통신상태(수신)

'## 1.1.12: 이상대(2005-07-28)
'   - 자동재검 사용유무 변수추가
Private mAutoRerun  As AutoRerunEnum    '자동재검 사용유무

Public Event EqpError(ByVal pCode As String)
Public Event SpcError(ByVal pCode As String, ByVal pBarNo As String)
Public Event SpcErrorX(ByVal pCode As String, ByVal pBarNo As String, ByVal pPtId As String, ByVal pName As String)
''Private mSQL        As clsIISLibrarySql

Public Property Get IntNms() As clsIISIntNms
    Set IntNms = mIntNms
End Property

Public Property Get AccInfos() As clsIISAccInfos
    Set AccInfos = mAccInfos
End Property

Public Property Get Buffers() As clsIISBuffers
    Set Buffers = mBuffers
End Property

Public Property Get BarPos() As BarPosEnum
    BarPos = mBarPos
End Property

Public Property Get Phase() As Long
    Phase = mPhase
End Property

Public Property Let Phase(ByVal vData As Long)
    mPhase = vData
End Property

Public Property Get SndPhase() As Long
    SndPhase = mSndPhase
End Property

Public Property Let SndPhase(ByVal vData As Long)
    mSndPhase = vData
End Property

Public Property Get BufCnt() As Long
    BufCnt = mBufCnt
End Property

Public Property Let BufCnt(ByVal vData As Long)
    mBufCnt = vData
End Property

Public Property Get State() As String
    State = mState
End Property

Public Property Let State(ByVal vData As String)
    mState = vData
End Property

Public Property Let FrameNo(ByVal vData As Long)
    mFrameNo = vData
End Property

Public Property Get IsETB() As Boolean
    IsETB = mIsETB
End Property

Public Property Let IsETB(ByVal vData As Boolean)
    mIsETB = vData
End Property

Public Property Get ACKFg() As Long
    ACKFg = mACKFg
End Property

Public Property Let ACKFg(ByVal vData As Long)
    mACKFg = vData
End Property

'## 1.1.10: 이상대(2005-07-06)
'   - 수신상태 프로퍼티 추가
Public Property Get RcvPhase() As Long
    RcvPhase = mRcvPhase
End Property

Public Property Let RcvPhase(ByVal vData As Long)
    mRcvPhase = vData
End Property

'## 1.1.12: 이상대(2005-07-28)
'   - 자동재검 사용유무 프로퍼티 추가
Public Property Get AutoRerun() As AutoRerunEnum
    AutoRerun = mAutoRerun
End Property

Public Property Let AutoRerun(ByVal vData As AutoRerunEnum)
    mAutoRerun = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 레지정보 조회, 통신설정 정보조회, 장비별 검사항목 컬렉션 생성,
'          로그파일 오픈, 보관기간이 지난 데이터 삭제
'-----------------------------------------------------------------------------'
Public Sub SetConfig(ByVal pEqpCd As String, ByVal pEqpKey As String)
    mEqpCd = pEqpCd
    mEqpKey = pEqpKey
    
    '## 레지정보조회, 통신설정 정보조회, 장비별 검사항목 컬렉션 생성
    '   로그파일 오픈, 보관기간이 지난 데이터 삭제
    Call GetEqpInfo
    Call GetIntTestCds
    Call SetLogFile
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 레지정보 조회
'-----------------------------------------------------------------------------'
Private Sub GetEqpInfo()
    Dim objEqpChoice As clsIISEqpChoice     '장비선택 클래스
    Dim strErrMsg    As String              '에러메세지

    '## 1.1.12: 이상대(2005-07-28)
    '   - 자동재검 사용유무 프로퍼티 추가
    Set objEqpChoice = New clsIISEqpChoice
    With objEqpChoice
        .GetEqp
        If mEqpCd = .EqpCd1 Then
            mBarcode = .Barcode1
            mAutoVfy = .AutoVfy1
            mHLCheck = .HLCheck1
            mBarPos = .BarPos1
            mAutoRerun = .AutoRerun1
        ElseIf mEqpCd = .EqpCd2 Then
            mBarcode = .Barcode2
            mAutoVfy = .AutoVfy2
            mHLCheck = .HLCheck2
            mBarPos = .BarPos2
            mAutoRerun = .AutoRerun2
        Else
            '## 에러 이벤트 발생
            strErrMsg = mEqpCd & "장비에 대한 레지스트리 정보가 없습니다."
            Error.SetLog App.EXEName, "clsIISInterface", "GetAutoVfy", strErrMsg, Now
            RaiseEvent EqpError("E001")
        End If
    End With
    
    Set objEqpChoice = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 장비별 검사항목 컬렉션 클래스 생성(LIS검사항목 기준)
'-----------------------------------------------------------------------------'
Private Sub GetIntTestCds()
    Dim Rs           As ADODB.Recordset
    Dim objIntNm     As clsIISIntNm         '장비별 검사항목 클래스
    Dim strTestCd    As String              '검사코드
    Dim strIntNm     As String              '검사명
    Dim strIntBase   As String              '장비기준 검사명
    Dim strFrVal     As String              'From Value
    Dim strToVal     As String              'To Value
    Dim strErrMsg    As String              '에러메시지

On Error GoTo Errors
    Set objIntNm = New clsIISIntNm
    Set Rs = objIntNm.GetTestCdIntNms(mEqpCd)
    Set objIntNm = Nothing

    '## NOTE: 검사코드가 중복됐을 경우 처리 할것!
    If Not (Rs.BOF Or Rs.EOF) Then
        Do Until Rs.EOF
            strTestCd = Rs.Fields("TESTCD").Value
            strIntNm = Rs.Fields("INTNM").Value
            strIntBase = Rs.Fields("INTBASE").Value & ""
            strFrVal = Rs.Fields("FRVAL").Value & ""
            strToVal = Rs.Fields("TOVAL").Value & ""
    
            Call mIntNms.AddIIS(strTestCd, strIntNm, strIntBase, strFrVal, strToVal)
            Rs.MoveNext
        Loop
    Else
        '## 에러 이벤트 발생
        strErrMsg = mEqpCd & " 장비에 장비별 검사항목이 설정되지 않었습니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetIntTestCds", strErrMsg, Now
        RaiseEvent EqpError("E002")
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Sub

Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISInterface", "GetIntTestCds", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당장비의 통신설정상태를 조회
'   반환 : 장비 통신설정 객체
'-----------------------------------------------------------------------------'
Public Function GetEqpComm() As clsIISEqpComm
    Dim objEqpComm As clsIISEqpComm     '장비 통신설정 클래스
    Dim strErrMsg  As String            '에러 메세지

    Set objEqpComm = New clsIISEqpComm
    Call objEqpComm.GetEqpComm(mEqpCd)

    With objEqpComm
        If .Port = "" Or .BaudRate = "" Or .Databit = "" Or .Stopbit = "" Or .Paritybit = "" Then
            strErrMsg = mEqpCd & " 장비에 대한 통신설정이 되어 있지 않습니다."
            Error.SetLog App.EXEName, "clsIISInterface", "GetEqpComm", strErrMsg, Now
            RaiseEvent EqpError("E003")
            GoTo EndLine
        End If

        '## ClientDb 보관일
        If .StoredDt = "" Then
            mStoreDt = 30
        Else
            mStoreDt = CLng(.StoredDt)
        End If
    End With
    Set GetEqpComm = objEqpComm

EndLine:
    Set objEqpComm = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호의 오더정보 조회1
'   인수 :
'       - pBarNo : 바코드번호
'   반환 : clsIISAccInfo 객체
'-----------------------------------------------------------------------------'
Public Function GetAccInfo(ByVal pBarNo As String) As clsIISAccInfo
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim objAccess As clsIISAccession     '접수내역 클래스
    Dim strSpcYy   As String            '접수연도
    Dim lngSpcNo   As Long              '접수순번
    Dim strErrMsg  As String            '에러메시지
    Dim ColSuccess  As Boolean
    
    If Len(Trim(pBarNo)) < 3 Then
        '## 잘못된 바코드번호 일때
        strErrMsg = "바코드번호:" & pBarNo & "는 잘못된 바코드번호 입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S013", pBarNo)
        GoTo EndLine
    End If
    
    pBarNo = Format$(pBarNo, String$(SPCLEN, "#"))
    strSpcYy = Mid$(pBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(pBarNo, SPCYYLEN + 1, SPCNOLEN))
    
    '## 접수내역 조회
    Set objAccInfo = New clsIISAccInfo
    Set objAccess = New clsIISAccession
    
AeeessionLine:
    If objAccInfo.GetAccInfo(strSpcYy, lngSpcNo) = -1 Then
        '## 접수내역이 없을때의 이벤트 발생
        strErrMsg = "바코드번호:" & pBarNo & "에 대한 접수내역이 없습니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S001", pBarNo)
        GoTo EndLine
    End If
    
    With objAccInfo
        Select Case .StsCd
            Case "0"    '## 처방상태
                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcErrorX("S002", pBarNo, .PtId, .Name)
                GoTo EndLine
            Case "1"    '## 채혈상태
                '## 접수진행
                ColSuccess = objAccess.DoAccession(objAccInfo.WorkArea, objAccInfo.AccDt, objAccInfo.AccSeq, EmpId)
                If ColSuccess = False Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S003", pBarNo, .PtId, .Name)
                    GoTo EndLine
                Else
                    GoTo AeeessionLine
                End If
            Case "5", "6"   '## 결과확인, 수정상태
                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcErrorX("S004", pBarNo, .PtId, .Name)
                GoTo EndLine
            Case "7"    '## 접수취소
                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcErrorX("S005", pBarNo, .PtId, .Name)
                GoTo EndLine
        End Select
        
        '## 일반 & QC 검체를 구분하여 결과내역 조회(유효숫자, D/P수치도 조회)
        If .QcFg = "0" Then
            If .TestDiv = "0" Then      '## 일반검사
                '## 1.1.12: 이상대(2005-07-28)
                '   - 자동재검 사용유무를 인수로 전달하도록 수정
                If .GetResultInfo(mAutoRerun) = -1 Then
                    '## 일반검사 결과내역이 없을때의 이벤트 발생
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            ElseIf .TestDiv = "1" Then  '## 특수검사
                '## 1.1.14: 이상대(2005-10-11)
                '   - 특수검사 결과내역 조회 추가
                If .GetSpecialResultInfo = -1 Then
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            '-- BactAlert는 배양장비 아님
            ElseIf .TestDiv = "2" Then  '## 미생물검사
            'ElseIf .TestDiv = "2" And mEqpCd <> "C015" Then  '## 미생물검사
                If .StsCd = "2" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 배양대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S017", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            
                '## 1.0.2: 이상대(2004-12-06)
                '   - 장비별 검사항목에 등록된 코드만 조회하기위해 mIntNms을 전달하도록 수정
                If .GetMICResultInfo(mIntNms) = -1 Then
                    '## 미생물 결과내역이 없을때의 이벤트 발생
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                '## Growth된 검체인지 체크
                If .MICResult.WSBody.GenFg <> "2" And mEqpCd <> "C015" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 Growth된 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                '## 감수성 결과대상인지 체크
                If .MICResult.RstType <> "S" And .MICResult.RstType <> "C" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 감수성 대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            End If
        ElseIf .QcFg = "1" Then
            '## QC결과내역 조회
            If .GetQCResultInfo = -1 Then
                '## QC결과내역이 없을때의 이벤트 발생
                strErrMsg = "바코드번호:" & pBarNo & "에 대한 QC결과내역이 없습니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcError("S012", pBarNo)
                GoTo EndLine
            End If
        End If
        
        '## 장비기준 검사항목 설정(일반,QC검체일때만) - 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" Then
            Call SetTestItem(objAccInfo)
        End If
        
        '## 참고치, Panic 수치설정(일반검체만)
        If .QcFg = "0" And .TestDiv = "0" Then
            Call SetReference(objAccInfo)
        End If
        
        '## 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" And .SendCnt < 1 Then
            '## 검사할 항목이 없음
            strErrMsg = "바코드번호: " & pBarNo & "는 장비로 검사할 항목이 없습니다."
            Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
            RaiseEvent SpcErrorX("S007", pBarNo, .PtId, .Name)
            GoTo EndLine
        End If
    End With
    
    If mAccInfos.Exist(objAccInfo.SpcYy, objAccInfo.SpcNo) Then
        '## 현재 검사중인 검체 체크
        strErrMsg = "바코드번호: " & pBarNo & "는 Worklist에 존재하는 검체입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcErrorX("S008", pBarNo, objAccInfo.PtId, objAccInfo.Name)
        GoTo EndLine
    End If
    
    Call mAccInfos.AddInterX(objAccInfo)
    Set GetAccInfo = objAccInfo
    
EndLine:
    Set objAccInfo = Nothing
End Function

'''-----------------------------------------------------------------------------'
'''   기능 : 해당 날짜의 오더정보 조회
'''   인수 :
'''       - pFrDT : 조회일자
'''   반환 : Recordset
'''-----------------------------------------------------------------------------'
''Public Function GetAccInfo_IT3000(ByVal pFrDT As String) As Recordset
''
''    Dim Rs As ADODB.Recordset
''
'''On Error GoTo Errors:
''
''    Set Rs = DbCon.Execute(mSQL.SelectAccInfo("2014", "123456"), , adCmdText)
''    If Not (Rs.BOF Or Rs.EOF) Then
''        'GetAccInfo_IT3000 = Rs
''    Else
''        'GetAccInfo = -1
''    End If
''
''    Rs.Close
''    Set Rs = Nothing
''    Exit Function
''
''End Function
'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호의 오더정보 조회1
'   인수 :
'       - pBarNo : 바코드번호
'   반환 : clsIISAccInfo 객체
'-----------------------------------------------------------------------------'
Public Function GetAccInfo_New(ByVal pBarNo As String, ByVal pPOC As String) As clsIISAccInfo
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim objAccess As clsIISAccession     '접수내역 클래스
    Dim strSpcYy   As String            '접수연도
    Dim lngSpcNo   As Long              '접수순번
    Dim strErrMsg  As String            '에러메시지
    Dim ColSuccess  As Boolean
    
    If Len(Trim(pBarNo)) < 3 Then
        '## 잘못된 바코드번호 일때
        strErrMsg = "바코드번호:" & pBarNo & "는 잘못된 바코드번호 입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S013", pBarNo)
        GoTo EndLine
    End If
    
    pBarNo = Format$(pBarNo, String$(SPCLEN, "#"))
    strSpcYy = Mid$(pBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(pBarNo, SPCYYLEN + 1, SPCNOLEN))
    
    '## 접수내역 조회
    Set objAccInfo = New clsIISAccInfo
    Set objAccess = New clsIISAccession
    
AeeessionLine:
    If objAccInfo.GetAccInfo(strSpcYy, lngSpcNo) = -1 Then
        '## 접수내역이 없을때의 이벤트 발생
        strErrMsg = "바코드번호:" & pBarNo & "에 대한 접수내역이 없습니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S001", pBarNo)
        GoTo EndLine
    End If
    
    With objAccInfo
        Select Case .StsCd
            Case "0"    '## 처방상태
                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcErrorX("S002", pBarNo, .PtId, .Name)
                GoTo EndLine
            Case "1"    '## 채혈상태
                '## 접수진행
                If pPOC = "1" Then
                    ColSuccess = objAccess.DoAccession(objAccInfo.WorkArea, objAccInfo.AccDt, objAccInfo.AccSeq, EmpId)
                    If ColSuccess = False Then
                        strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                        RaiseEvent SpcErrorX("S003", pBarNo, .PtId, .Name)
                        GoTo EndLine
                    Else
                        GoTo AeeessionLine
                    End If
                Else
                    strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S003", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            Case "5", "6"   '## 결과확인, 수정상태
                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcErrorX("S004", pBarNo, .PtId, .Name)
                GoTo EndLine
            Case "7"    '## 접수취소
                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcErrorX("S005", pBarNo, .PtId, .Name)
                GoTo EndLine
        End Select
        
        '## 일반 & QC 검체를 구분하여 결과내역 조회(유효숫자, D/P수치도 조회)
        If .QcFg = "0" Then
            If .TestDiv = "0" Then      '## 일반검사
                '## 1.1.12: 이상대(2005-07-28)
                '   - 자동재검 사용유무를 인수로 전달하도록 수정
                If .GetResultInfo(mAutoRerun) = -1 Then
                    '## 일반검사 결과내역이 없을때의 이벤트 발생
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            ElseIf .TestDiv = "1" Then  '## 특수검사
                '## 1.1.14: 이상대(2005-10-11)
                '   - 특수검사 결과내역 조회 추가
                If .GetSpecialResultInfo = -1 Then
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            '-- BactAlert는 배양장비 아님
            ElseIf .TestDiv = "2" Then  '## 미생물검사
            'ElseIf .TestDiv = "2" And mEqpCd <> "C015" Then  '## 미생물검사
                If .StsCd = "2" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 배양대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S017", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            
                '## 1.0.2: 이상대(2004-12-06)
                '   - 장비별 검사항목에 등록된 코드만 조회하기위해 mIntNms을 전달하도록 수정
                If .GetMICResultInfo(mIntNms) = -1 Then
                    '## 미생물 결과내역이 없을때의 이벤트 발생
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                '## Growth된 검체인지 체크
                If .MICResult.WSBody.GenFg <> "2" And mEqpCd <> "C015" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 Growth된 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                '## 감수성 결과대상인지 체크
                If .MICResult.RstType <> "S" And .MICResult.RstType <> "C" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 감수성 대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            End If
        ElseIf .QcFg = "1" Then
            '## QC결과내역 조회
            If .GetQCResultInfo = -1 Then
                '## QC결과내역이 없을때의 이벤트 발생
                strErrMsg = "바코드번호:" & pBarNo & "에 대한 QC결과내역이 없습니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcError("S012", pBarNo)
                GoTo EndLine
            End If
        End If
        
        '## 장비기준 검사항목 설정(일반,QC검체일때만) - 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" Then
            Call SetTestItem(objAccInfo)
        End If
        
        '## 참고치, Panic 수치설정(일반검체만)
        If .QcFg = "0" And .TestDiv = "0" Then
            Call SetReference(objAccInfo)
        End If
        
        '## 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" And .SendCnt < 1 Then
            '## 검사할 항목이 없음
            strErrMsg = "바코드번호: " & pBarNo & "는 장비로 검사할 항목이 없습니다."
            Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
            RaiseEvent SpcErrorX("S007", pBarNo, .PtId, .Name)
            GoTo EndLine
        End If
    End With
    
    If mAccInfos.Exist(objAccInfo.SpcYy, objAccInfo.SpcNo) Then
        '## 현재 검사중인 검체 체크
        strErrMsg = "바코드번호: " & pBarNo & "는 Worklist에 존재하는 검체입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcErrorX("S008", pBarNo, objAccInfo.PtId, objAccInfo.Name)
        GoTo EndLine
    End If
    
    Call mAccInfos.AddInterX(objAccInfo)
    Set GetAccInfo_New = objAccInfo
    
EndLine:
    Set objAccInfo = Nothing
End Function

Public Function GetAccInfo_New_Temp(ByVal pBarNo As String, ByVal pPOC As String) As clsIISAccInfo
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim objAccess As clsIISAccession     '접수내역 클래스
    Dim strSpcYy   As String            '접수연도
    Dim lngSpcNo   As Long              '접수순번
    Dim strErrMsg  As String            '에러메시지
    Dim ColSuccess  As Boolean
    
    If Len(Trim(pBarNo)) < 3 Then
        '## 잘못된 바코드번호 일때
        strErrMsg = "바코드번호:" & pBarNo & "는 잘못된 바코드번호 입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S013", pBarNo)
        GoTo EndLine
    End If
    
    pBarNo = Format$(pBarNo, String$(SPCLEN, "#"))
    strSpcYy = Mid$(pBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(pBarNo, SPCYYLEN + 1, SPCNOLEN))
    
    '## 접수내역 조회
    Set objAccInfo = New clsIISAccInfo
    Set objAccess = New clsIISAccession
    
AeeessionLine:
    If objAccInfo.GetAccInfo(strSpcYy, lngSpcNo) = -1 Then
        '## 접수내역이 없을때의 이벤트 발생
        strErrMsg = "바코드번호:" & pBarNo & "에 대한 접수내역이 없습니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S001", pBarNo)
        GoTo EndLine
    End If
    
    With objAccInfo
'        Select Case .StsCd
'            Case "0"    '## 처방상태
'                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
'                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
'                RaiseEvent SpcErrorX("S002", pBarNo, .Ptid, .Name)
'                GoTo EndLine
'            Case "1"    '## 채혈상태
                '## 접수진행
                If pPOC = "1" Then
                    ColSuccess = objAccess.DoAccession_Temp(objAccInfo.WorkArea, objAccInfo.AccDt, objAccInfo.AccSeq, EmpId)
'                    If ColSuccess = False Then
'                        strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
'                        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
'                        RaiseEvent SpcErrorX("S003", pBarNo, .Ptid, .Name)
'                        GoTo EndLine
'                    Else
'                        GoTo AeeessionLine
'                    End If
'                Else
'                    strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
'                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
'                    RaiseEvent SpcErrorX("S003", pBarNo, .Ptid, .Name)
'                    GoTo EndLine
                End If
'            Case "5", "6"   '## 결과확인, 수정상태
'                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
'                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
'                RaiseEvent SpcErrorX("S004", pBarNo, .Ptid, .Name)
'                GoTo EndLine
'            Case "7"    '## 접수취소
'                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
'                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
'                RaiseEvent SpcErrorX("S005", pBarNo, .Ptid, .Name)
'                GoTo EndLine
'        End Select
'
        '## 일반 & QC 검체를 구분하여 결과내역 조회(유효숫자, D/P수치도 조회)
        If .QcFg = "0" Then
            If .TestDiv = "0" Then      '## 일반검사
                '## 1.1.12: 이상대(2005-07-28)
                '   - 자동재검 사용유무를 인수로 전달하도록 수정
                If .GetResultInfo(mAutoRerun) = -1 Then
                    '## 일반검사 결과내역이 없을때의 이벤트 발생
'''                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
'''                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
'''                    RaiseEvent SpcErrorX("S006", pBarNo, .Ptid, .Name)
                    GoTo EndLine
                End If
            ElseIf .TestDiv = "1" Then  '## 특수검사
                '## 1.1.14: 이상대(2005-10-11)
                '   - 특수검사 결과내역 조회 추가
                If .GetSpecialResultInfo = -1 Then
'''                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
'''                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
'''                    RaiseEvent SpcErrorX("S006", pBarNo, .Ptid, .Name)
                    GoTo EndLine
                End If
            '-- BactAlert는 배양장비 아님
            ElseIf .TestDiv = "2" Then  '## 미생물검사
            'ElseIf .TestDiv = "2" And mEqpCd <> "C015" Then  '## 미생물검사
                If .StsCd = "2" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 배양대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S017", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            
                '## 1.0.2: 이상대(2004-12-06)
                '   - 장비별 검사항목에 등록된 코드만 조회하기위해 mIntNms을 전달하도록 수정
                If .GetMICResultInfo(mIntNms) = -1 Then
                    '## 미생물 결과내역이 없을때의 이벤트 발생
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                '## Growth된 검체인지 체크
                If .MICResult.WSBody.GenFg <> "2" And mEqpCd <> "C015" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 Growth된 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                '## 감수성 결과대상인지 체크
                If .MICResult.RstType <> "S" And .MICResult.RstType <> "C" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 감수성 대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            End If
        ElseIf .QcFg = "1" Then
            '## QC결과내역 조회
            If .GetQCResultInfo = -1 Then
                '## QC결과내역이 없을때의 이벤트 발생
                strErrMsg = "바코드번호:" & pBarNo & "에 대한 QC결과내역이 없습니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcError("S012", pBarNo)
                GoTo EndLine
            End If
        End If
        
        '## 장비기준 검사항목 설정(일반,QC검체일때만) - 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" Then
            Call SetTestItem(objAccInfo)
        End If
        
        '## 참고치, Panic 수치설정(일반검체만)
        If .QcFg = "0" And .TestDiv = "0" Then
            Call SetReference(objAccInfo)
        End If
        
        '## 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" And .SendCnt < 1 Then
            '## 검사할 항목이 없음
            strErrMsg = "바코드번호: " & pBarNo & "는 장비로 검사할 항목이 없습니다."
            Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
            RaiseEvent SpcErrorX("S007", pBarNo, .PtId, .Name)
            GoTo EndLine
        End If
    End With
    
    If mAccInfos.Exist(objAccInfo.SpcYy, objAccInfo.SpcNo) Then
        '## 현재 검사중인 검체 체크
        strErrMsg = "바코드번호: " & pBarNo & "는 Worklist에 존재하는 검체입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcErrorX("S008", pBarNo, objAccInfo.PtId, objAccInfo.Name)
        GoTo EndLine
    End If
    
    Call mAccInfos.AddInterX(objAccInfo)
    Set GetAccInfo_New_Temp = objAccInfo
    
EndLine:
    Set objAccInfo = Nothing
End Function

Public Function GetAccInfo_New_Temp_Dual(ByVal pBarNo As String, ByVal pPOC As String) As clsIISAccInfo
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim objAccess As clsIISAccession     '접수내역 클래스
    Dim strSpcYy   As String            '접수연도
    Dim lngSpcNo   As Long              '접수순번
    Dim strErrMsg  As String            '에러메시지
    Dim ColSuccess  As Boolean
    
    If Len(Trim(pBarNo)) < 3 Then
        '## 잘못된 바코드번호 일때
        strErrMsg = "바코드번호:" & pBarNo & "는 잘못된 바코드번호 입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S013", pBarNo)
        GoTo EndLine
    End If
    
    pBarNo = Format$(pBarNo, String$(SPCLEN, "#"))
    strSpcYy = Mid$(pBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(pBarNo, SPCYYLEN + 1, SPCNOLEN))
    
    '## 접수내역 조회
    Set objAccInfo = New clsIISAccInfo
    Set objAccess = New clsIISAccession
    
AeeessionLine:
    If objAccInfo.GetAccInfo(strSpcYy, lngSpcNo) = -1 Then
        '## 접수내역이 없을때의 이벤트 발생
        strErrMsg = "바코드번호:" & pBarNo & "에 대한 접수내역이 없습니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S001", pBarNo)
        GoTo EndLine
    End If
    
    With objAccInfo
        '## 접수진행
        If pPOC = "1" Then
        '    ColSuccess = objAccess.DoAccession_Temp(objAccInfo.WorkArea, objAccInfo.AccDt, objAccInfo.AccSeq, EmpId)
        End If
        
        '## 일반 & QC 검체를 구분하여 결과내역 조회(유효숫자, D/P수치도 조회)
        If .QcFg = "0" Then
            If .TestDiv = "0" Then      '## 일반검사
                If .GetResultInfo(mAutoRerun) = -1 Then
                    GoTo EndLine
                End If
            ElseIf .TestDiv = "1" Then  '## 특수검사
                If .GetSpecialResultInfo = -1 Then
                    GoTo EndLine
                End If
            ElseIf .TestDiv = "2" Then  '## 미생물검사
                If .StsCd = "2" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 배양대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S017", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                If .GetMICResultInfo(mIntNms) = -1 Then
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                If .MICResult.WSBody.GenFg <> "2" And mEqpCd <> "C015" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 Growth된 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                '## 감수성 결과대상인지 체크
                If .MICResult.RstType <> "S" And .MICResult.RstType <> "C" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 감수성 대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            End If
        ElseIf .QcFg = "1" Then
            '## QC결과내역 조회
            If .GetQCResultInfo = -1 Then
                '## QC결과내역이 없을때의 이벤트 발생
                strErrMsg = "바코드번호:" & pBarNo & "에 대한 QC결과내역이 없습니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcError("S012", pBarNo)
                GoTo EndLine
            End If
        End If
        
        '## 장비기준 검사항목 설정(일반,QC검체일때만) - 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" Then
            Call SetTestItem(objAccInfo)
        End If
        
        '## 참고치, Panic 수치설정(일반검체만)
        If .QcFg = "0" And .TestDiv = "0" Then
            Call SetReference(objAccInfo)
        End If
        
        '## 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" And .SendCnt < 1 Then
            '## 검사할 항목이 없음
            strErrMsg = "바코드번호: " & pBarNo & "는 장비로 검사할 항목이 없습니다."
            Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
            RaiseEvent SpcErrorX("S007", pBarNo, .PtId, .Name)
            GoTo EndLine
        End If
    End With
    
    If mAccInfos.Exist(objAccInfo.SpcYy, objAccInfo.SpcNo) Then
        '## 현재 검사중인 검체 체크
        strErrMsg = "바코드번호: " & pBarNo & "는 Worklist에 존재하는 검체입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcErrorX("S008", pBarNo, objAccInfo.PtId, objAccInfo.Name)
        GoTo EndLine
    End If
    
    Call mAccInfos.AddInterX(objAccInfo)
    Set GetAccInfo_New_Temp_Dual = objAccInfo
    
EndLine:
    Set objAccInfo = Nothing
End Function

Public Function GetAccInfo_New_Temp1(ByVal pBarNo As String, ByVal pPOC As String) As clsIISAccInfo
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim objAccess As clsIISAccession     '접수내역 클래스
    Dim strSpcYy   As String            '접수연도
    Dim lngSpcNo   As Long              '접수순번
    Dim strErrMsg  As String            '에러메시지
    Dim ColSuccess  As Boolean
    
    If Len(Trim(pBarNo)) < 3 Then
        '## 잘못된 바코드번호 일때
        strErrMsg = "바코드번호:" & pBarNo & "는 잘못된 바코드번호 입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S013", pBarNo)
        GoTo EndLine
    End If
    
    pBarNo = Format$(pBarNo, String$(SPCLEN, "#"))
    strSpcYy = Mid$(pBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(pBarNo, SPCYYLEN + 1, SPCNOLEN))
    
    '## 접수내역 조회
    Set objAccInfo = New clsIISAccInfo
    Set objAccess = New clsIISAccession
    
AeeessionLine:
    If objAccInfo.GetAccInfo(strSpcYy, lngSpcNo) = -1 Then
        '## 접수내역이 없을때의 이벤트 발생
        strErrMsg = "바코드번호:" & pBarNo & "에 대한 접수내역이 없습니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcError("S001", pBarNo)
        GoTo EndLine
    End If
    
    With objAccInfo
        Select Case .StsCd
            Case "0"    '## 처방상태
                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcErrorX("S002", pBarNo, .PtId, .Name)
                GoTo EndLine
            Case "1"    '## 채혈상태
                '## 접수진행
                If pPOC = "1" And ColSuccess = False Then
                    ColSuccess = objAccess.DoAccession_Temp1(objAccInfo.WorkArea, objAccInfo.AccDt, objAccInfo.AccSeq, EmpId)
                    If ColSuccess = False Then
'                        strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
'                        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
'                        RaiseEvent SpcErrorX("S003", pBarNo, .Ptid, .Name)
                        GoTo EndLine
                    Else
                        GoTo AeeessionLine
                    End If
'                Else
'                    strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
'                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
'                    RaiseEvent SpcErrorX("S003", pBarNo, .Ptid, .Name)
'                    GoTo EndLine
                End If
            Case "5", "6"   '## 결과확인, 수정상태
                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcErrorX("S004", pBarNo, .PtId, .Name)
                GoTo EndLine
            Case "7"    '## 접수취소
                strErrMsg = "바코드번호:" & pBarNo & "는 접수상태가 아닙니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcErrorX("S005", pBarNo, .PtId, .Name)
                GoTo EndLine
        End Select
        
        '## 일반 & QC 검체를 구분하여 결과내역 조회(유효숫자, D/P수치도 조회)
        If .QcFg = "0" Then
            If .TestDiv = "0" Then      '## 일반검사
                '## 1.1.12: 이상대(2005-07-28)
                '   - 자동재검 사용유무를 인수로 전달하도록 수정
                If .GetResultInfo(mAutoRerun) = -1 Then
                    '## 일반검사 결과내역이 없을때의 이벤트 발생
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            ElseIf .TestDiv = "1" Then  '## 특수검사
                '## 1.1.14: 이상대(2005-10-11)
                '   - 특수검사 결과내역 조회 추가
                If .GetSpecialResultInfo = -1 Then
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            '-- BactAlert는 배양장비 아님
            ElseIf .TestDiv = "2" Then  '## 미생물검사
            'ElseIf .TestDiv = "2" And mEqpCd <> "C015" Then  '## 미생물검사
                If .StsCd = "2" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 배양대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S017", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            
                '## 1.0.2: 이상대(2004-12-06)
                '   - 장비별 검사항목에 등록된 코드만 조회하기위해 mIntNms을 전달하도록 수정
                If .GetMICResultInfo(mIntNms) = -1 Then
                    '## 미생물 결과내역이 없을때의 이벤트 발생
                    strErrMsg = "바코드번호:" & pBarNo & "에 대한 결과내역이 없습니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S006", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                '## Growth된 검체인지 체크
                If .MICResult.WSBody.GenFg <> "2" And mEqpCd <> "C015" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 Growth된 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
                
                '## 감수성 결과대상인지 체크
                If .MICResult.RstType <> "S" And .MICResult.RstType <> "C" Then
                    strErrMsg = "바코드번호:" & pBarNo & "는 감수성 대상 검체가 아닙니다."
                    Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                    RaiseEvent SpcErrorX("S014", pBarNo, .PtId, .Name)
                    GoTo EndLine
                End If
            End If
        ElseIf .QcFg = "1" Then
            '## QC결과내역 조회
            If .GetQCResultInfo = -1 Then
                '## QC결과내역이 없을때의 이벤트 발생
                strErrMsg = "바코드번호:" & pBarNo & "에 대한 QC결과내역이 없습니다."
                Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
                RaiseEvent SpcError("S012", pBarNo)
                GoTo EndLine
            End If
        End If
        
        '## 장비기준 검사항목 설정(일반,QC검체일때만) - 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" Then
            Call SetTestItem(objAccInfo)
        End If
        
        '## 참고치, Panic 수치설정(일반검체만)
        If .QcFg = "0" And .TestDiv = "0" Then
            Call SetReference(objAccInfo)
        End If
        
        '## 미생물 검사는 검사항목 설정이 의미없음
        If .TestDiv <> "2" And .SendCnt < 1 Then
            '## 검사할 항목이 없음
            strErrMsg = "바코드번호: " & pBarNo & "는 장비로 검사할 항목이 없습니다."
            Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
            RaiseEvent SpcErrorX("S007", pBarNo, .PtId, .Name)
            GoTo EndLine
        End If
    End With
    
    If mAccInfos.Exist(objAccInfo.SpcYy, objAccInfo.SpcNo) Then
        '## 현재 검사중인 검체 체크
        strErrMsg = "바코드번호: " & pBarNo & "는 Worklist에 존재하는 검체입니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfo", strErrMsg, Now
        RaiseEvent SpcErrorX("S008", pBarNo, objAccInfo.PtId, objAccInfo.Name)
        GoTo EndLine
    End If
    
    Call mAccInfos.AddInterX(objAccInfo)
    Set GetAccInfo_New_Temp1 = objAccInfo
    
EndLine:
    Set objAccInfo = Nothing
End Function


'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호의 오더정보 조회2
'   인수 :
'       - pBarNo : 바코드번호
'   반환 : clsIISAccInfo 객체
'-----------------------------------------------------------------------------'
Public Function GetAccInfoX(ByVal pBarNo As String) As clsIISAccInfo
    Dim strSpcYy   As String            '접수연도
    Dim lngSpcNo   As Long              '접수순번
    
    pBarNo = Format$(pBarNo, String$(SPCLEN, "#"))
    strSpcYy = Mid$(pBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(pBarNo, SPCYYLEN + 1, SPCNOLEN))
    
    '## 접수내역 컬렉션에 바코드번호가 존재하면 해당 접수내역 반환, 없으면 접수내역 조회
    If mAccInfos.Exist(strSpcYy, lngSpcNo) Then
        Set GetAccInfoX = mAccInfos(strSpcYy, lngSpcNo)
    Else
        Set GetAccInfoX = GetAccInfo(pBarNo)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 접수번호의 접수정보 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSe    : 접수순번
'   반환 : clsIISAccInfo 객체
'-----------------------------------------------------------------------------'
Public Function GetAccInfoByAccNo(ByVal pWorkArea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As clsIISAccInfo
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim strBarNo    As String           '바코드번호
    Dim strSpcYy    As String           '검체연도
    Dim strErrMsg   As String           '에러메시지
    Dim lngSpcNo    As Long             '검체번호
    
    '## 1.해당 접수번호의 검체번호를 조회
    Set objAccInfo = New clsIISAccInfo
    strBarNo = objAccInfo.GetBarNoByAccNo(pWorkArea, pAccDt, pAccSeq)
    Set objAccInfo = Nothing
    
    If strBarNo = "-1" Then
        '## 접수내역이 없을때의 이벤트 발생
        strErrMsg = "접수번호:" & mGetAccNo(pWorkArea, pAccDt, pAccSeq) & "에 대한 접수내역이 없습니다."
        Error.SetLog App.EXEName, "clsIISInterface", "GetAccInfoByAccNo", strErrMsg, Now
        RaiseEvent SpcError("S016", mGetAccNo(pWorkArea, pAccDt, pAccSeq))
        Exit Function
    End If
    
    '## 2.접수내역 조회
    strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
    
    If mAccInfos.Exist(strSpcYy, lngSpcNo) = False Then
        Set GetAccInfoByAccNo = GetAccInfo(strBarNo)
    Else
        Set GetAccInfoByAccNo = mAccInfos(strSpcYy, lngSpcNo)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 결과내역에 검사코드에 해당하는 장비기준 검사항목 설정
'   인수 :
'       - pAccInfo : 접수내역 객체
'-----------------------------------------------------------------------------'
Private Sub SetTestItem(ByRef pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult      '결과내역 클래스
    Dim objQCResult As clsIISQCResult    'QC결과내역 클래스
    Dim objIntNm    As clsIISIntNm       '장비별 검사항목 클래스
    
    Dim strTestCd   As String            '검사코드
    Dim lngSendCnt  As Long              '장비로 전송할 검사항목 개수

    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            strTestCd = objResult.TestCd
            Set objIntNm = objResult.IntNm
            With objIntNm
                If mIntNms.Exist(strTestCd) Then
                    .IntNm = mIntNms(strTestCd).IntNm
                    .IntBase = mIntNms(strTestCd).IntBase
                    .FrVal = mIntNms(strTestCd).FrVal
                    .ToVal = mIntNms(strTestCd).ToVal
    
                    lngSendCnt = lngSendCnt + 1
                End If
            End With
        Next
        
        '## 해당장비로 검사할 항목이 아니면 삭제
        For Each objResult In pAccInfo.Results
            strTestCd = objResult.TestCd
            Set objIntNm = objResult.IntNm
            If objIntNm.IntNm = "" Then
                Call pAccInfo.Results.Remove(strTestCd)
            End If
        Next
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strTestCd = objQCResult.TestCd
            Set objIntNm = objQCResult.IntNm
            With objIntNm
                If mIntNms.Exist(strTestCd) Then
                    .IntNm = mIntNms(strTestCd).IntNm
                    .IntBase = mIntNms(strTestCd).IntBase
                    .FrVal = mIntNms(strTestCd).FrVal
                    .ToVal = mIntNms(strTestCd).ToVal
    
                    lngSendCnt = lngSendCnt + 1
                End If
            End With
        Next
        
        '## 해당장비로 검사할 항목이 아니면 삭제
        For Each objQCResult In pAccInfo.QCResults
            strTestCd = objQCResult.TestCd
            Set objIntNm = objQCResult.IntNm
            If objIntNm.IntNm = "" Then
                Call pAccInfo.QCResults.Remove(strTestCd)
            End If
        Next
    End If

    pAccInfo.SendCnt = lngSendCnt
    Set objIntNm = Nothing
    Set objResult = Nothing
    Set objQCResult = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 결과내역에 참고치, Panic 수치 설정
'   인수 :
'       - pAccInfo : 접수내역 객체
'-----------------------------------------------------------------------------'
Private Sub SetReference(ByRef pAccInfo As clsIISAccInfo)
    Dim Rs        As ADODB.Recordset
    Dim objResult As clsIISResult   '결과내역 클래스
    Dim objRef    As clsIISRef      '참고치 클래스
    Dim strTestCd As String         '검사코드
    Dim strSex    As String         'Delta, Panic수치를 조회성별(X:남자,Y:여자,Z:Both)

On Error GoTo Errors
    
    Set objRef = New clsIISRef
    With pAccInfo
        For Each objResult In .Results
            '## 최초 Both로 체크후 없으면 해당 성별로 체크
            strTestCd = objResult.TestCd
            
            '## 1.참고치 조회
            Set Rs = objRef.GetRef(strTestCd, .SpcCd, "B", .AgeDay)
            If Not (Rs.BOF Or Rs.EOF) Then
                If IsNull(Rs.Fields("REFFRVAL").Value) Or IsNull(Rs.Fields("REFTOVAL").Value) Then
                    objResult.Ref.RefFg = "0"
                Else
                    objResult.Ref.RefFg = "1"
                    objResult.Ref.RefFrVal = CSng(Rs.Fields("REFFRVAL").Value)
                    objResult.Ref.RefToVal = CSng(Rs.Fields("REFTOVAL").Value)
                    objResult.Ref.RefCd = Rs.Fields("REFCD").Value & ""
                End If
            Else
                Rs.Close
                Set Rs = objRef.GetRef(strTestCd, .SpcCd, .Sex, .AgeDay)
                If Not (Rs.BOF Or Rs.EOF) Then
                    If IsNull(Rs.Fields("REFFRVAL").Value) Or IsNull(Rs.Fields("REFTOVAL").Value) Then
                        objResult.Ref.RefFg = "0"
                    Else
                        objResult.Ref.RefFg = "1"
                        objResult.Ref.RefFrVal = CSng(Rs.Fields("REFFRVAL").Value)
                        objResult.Ref.RefToVal = CSng(Rs.Fields("REFTOVAL").Value)
                        objResult.Ref.RefCd = Rs.Fields("REFCD").Value & ""
                    End If
                Else
                    '## 참고치가 없을경우
                    objResult.Ref.RefFg = "0"
                End If
            End If
            Rs.Close
            
            '## 1.1.5: 이상대(2005-04-12)
            '   - 참고치 마스터에서 Panic수치를 조회할때만 쿼리하도록 수정
            '## 2.Panic수치 조회
            If IISPANICCHECK = "IIS103" Then
                Set Rs = objRef.GetPanic(strTestCd, .SpcCd, "Z", .AgeDay)
                If Not (Rs.BOF Or Rs.EOF) Then
                    If IsNull(Rs.Fields("PANICFRVAL").Value) Or IsNull(Rs.Fields("PANICTOVAL").Value) Then
                        objResult.Ref.PanicFg = "0"
                    Else
                        objResult.Ref.PanicFg = "1"
                        objResult.Ref.PanicFrVal = CSng(Rs.Fields("PANICFRVAL").Value)
                        objResult.Ref.PanicToVal = CSng(Rs.Fields("PANICTOVAL").Value)
                    End If
                Else
                    Rs.Close
                    
                    strSex = IIf(.Sex = "M", "X", "Y")
                    Set Rs = objRef.GetPanic(strTestCd, .SpcCd, strSex, .AgeDay)
                    If Not (Rs.BOF Or Rs.EOF) Then
                        If IsNull(Rs.Fields("PANICFRVAL").Value) Or IsNull(Rs.Fields("PANICTOVAL").Value) Then
                            objResult.Ref.PanicFg = "0"
                        Else
                            objResult.Ref.PanicFg = "1"
                            objResult.Ref.PanicFrVal = CSng(Rs.Fields("PANICFRVAL").Value)
                            objResult.Ref.PanicToVal = CSng(Rs.Fields("PANICTOVAL").Value)
                        End If
                    Else
                        '## Delta,Panic수치가 없을경우
                        objResult.Ref.PanicFg = "0"
                    End If
                End If
                Rs.Close
            End If
            Set Rs = Nothing
        Next
    End With
    
    Set objRef = Nothing
    Set objResult = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISInterface", "SetReference", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 결과내역에 참고치, Panic 수치 설정
'   인수 :
'       - pAccInfo : 접수내역 객체
'-----------------------------------------------------------------------------'
Private Sub SetDeltaPanic(ByRef pAccInfo As clsIISAccInfo)
    '## 1.1.5: 이상대(2005-04-12)
    '   - 함수이름의 모호함 때문에 SetReference로 변경함
    '   - 호환성 유지를 위해 남겨둠
    MsgBox "과거버전의 함수입니다. 대신 SetReference 함수를 사용하세요", vbInformation, "정보"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 결과판정
'   인수 :
'       - pIntInfo : 인터페이스 결과객체
'   반환 : 해당 바코드번호에 대해 접수내역이 없을경우: -1
'-----------------------------------------------------------------------------'
Public Function CheckResult(ByVal pIntInfo As clsIISIntInfo) As Long
    Dim objAccInfo   As clsIISAccInfo    '접수내역 클래스
    Dim objResult    As clsIISResult     '결과내역 클래스
    Dim objQCResult  As clsIISQCResult   'QC결과내역 클래스
    Dim objIntResult As clsIISIntResult  '인터페이스 결과클래스
    
    Dim strSpcYy    As String   '검체연도
    Dim lngSpcNo    As Long     '검체번호
    
    If Len(pIntInfo.BarNo) < 3 Then
        CheckResult = -1
        Exit Function
    End If
    
    strSpcYy = Mid$(pIntInfo.BarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(pIntInfo.BarNo, SPCYYLEN + 1, SPCNOLEN))
    '## 접수정보 조회
    If mAccInfos.Exist(strSpcYy, lngSpcNo) = False Then
        Set objAccInfo = GetAccInfo(pIntInfo.BarNo)
        If objAccInfo Is Nothing Then
            CheckResult = -1
            Exit Function
        End If
    Else
        Set objAccInfo = mAccInfos(strSpcYy, lngSpcNo)
    End If
    '## 해당 결과내역에 결과값입력
    If objAccInfo.QcFg = "0" Then       '## 일반검체
        For Each objIntResult In pIntInfo.IntResults
            For Each objResult In objAccInfo.Results
                If objIntResult.IntBase = objResult.IntNm.IntBase Then
                    objResult.IntResult = objIntResult.IntResult
                    objResult.Result = objIntResult.Result
                    objResult.IntInfo = objIntResult.Info
                    Exit For
                End If
            Next
        Next
    ElseIf objAccInfo.QcFg = "1" Then   '## QC검체
        For Each objIntResult In pIntInfo.IntResults
            For Each objQCResult In objAccInfo.QCResults
                If objIntResult.IntBase = objQCResult.IntNm.IntBase Then
                    objQCResult.IntResult = objIntResult.IntResult
                    objQCResult.Result = objIntResult.Result
                    Exit For
                End If
            Next
        Next
    End If

    '## 결과판정
    Call objAccInfo.CheckResult

    Set objIntResult = Nothing
    Set objQCResult = Nothing
    Set objResult = Nothing
    Set objAccInfo = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 검체번호에 대해 ClientDb에 결과입력
'   인수 :
'       - pSpcYy : 검체연도
'       - pSpcNo : 검체번호
'   반환 : True(성공) / False(실패)
'-----------------------------------------------------------------------------'
Public Function SaveClientDb(ByVal pSpcYy As String, ByVal pSpcNo As Long) As Boolean
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim objCAccinfo As clsIISCAccInfo   'ClientDb의 접수내역 클래스
    Dim strErrMsg   As String           '에러메세지
    
    Set objAccInfo = mAccInfos(pSpcYy, pSpcNo)
    Set objCAccinfo = New clsIISCAccInfo
    If objCAccinfo.SaveClientDb(objAccInfo, mEqpCd) = False Then
        '## ClientDb에 저장실패
        strErrMsg = "바코드번호: " & objAccInfo.GetBarNo & "의 ClientDb 저장에 실패했습니다."
        Error.SetLog App.EXEName, "clsIISInterface", "SaveClientDb", strErrMsg, Now
'        RaiseEvent SpcErrorX("S009", objAccInfo.GetBarNo, objAccInfo.PtId, objAccInfo.Name)
        GoTo EndLine
    End If
    SaveClientDb = True
    
EndLine:
    Set objCAccinfo = Nothing
    Set objAccInfo = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 결과입력
'   인수 :
'       - pSpcYy : 검체연도
'       - pSpcNo : 검체번호
'   반환 : True(성공) / False(실패)
'-----------------------------------------------------------------------------'
Public Function SaveResult(ByVal pSpcYy As String, ByVal pSpcNo As Long) As Boolean
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim strErrMsg  As String            '에러 메세지
    
    Set objAccInfo = mAccInfos(pSpcYy, pSpcNo)
    
    With objAccInfo
        '## 서버에 결과저장
        If mAutoVfy = AutoVfyEnum.ccNo Or .QcFg = "1" Then
            If .IISSaveServerNoVfy(mEqpCd) = False Then
                strErrMsg = "바코드번호: " & objAccInfo.GetBarNo & "의 서버 결과저장에 실패했습니다."
                Error.SetLog App.EXEName, "clsIISInterface", "SaveResult", strErrMsg, Now
                RaiseEvent SpcErrorX("S010", objAccInfo.GetBarNo, objAccInfo.PtId, objAccInfo.Name)
                GoTo EndLine
            End If
        ElseIf mAutoVfy = AutoVfyEnum.ccYes Then    '## 자동결과등록
            If .IISSaveServerAutoVfy(mEqpCd, mHLCheck) = False Then
                strErrMsg = "바코드번호: " & objAccInfo.GetBarNo & "의 서버 결과저장에 실패했습니다."
                Error.SetLog App.EXEName, "clsIISInterface", "SaveResult", strErrMsg, Now
                RaiseEvent SpcErrorX("S011", objAccInfo.GetBarNo, objAccInfo.PtId, objAccInfo.Name)
                GoTo EndLine
            End If
        End If
    End With
    SaveResult = True
    
EndLine:
    Set objAccInfo = Nothing
End Function

''-----------------------------------------------------------------------------'
''   기능 : 미생물 결과저장
''   인수 :
''       - pIntInfo : 장비에서 수신한결과 저장 클래스
''   반환 : Ture(성공) / False(실패)
''-----------------------------------------------------------------------------'
'Public Function SaveMICResult(ByVal pIntInfo As clsIISIntInfo) As Boolean
'    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
'    Dim strBarNo   As String    '바코드번호
'    Dim strSpcYy   As String    '검체연도
'    Dim lngSpcNo   As Long      '검체번호
'    Dim strErrMsg  As String    '에러메시지
'
'    strBarNo = pIntInfo.BarNo
'    strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
'    lngSpcNo = Val(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
'
'    Set objAccInfo = mAccInfos(strSpcYy, lngSpcNo)
'    If objAccInfo.MicroSaveServer(pIntInfo) = False Then
'        strErrMsg = "바코드번호: " & objAccInfo.GetBarNo & "의 서버 결과저장에 실패했습니다."
'        Error.SetLog App.EXEName, "clsIISInterface", "SaveMICResult", strErrMsg, Now
'        RaiseEvent SpcErrorX("S010", objAccInfo.GetBarNo, objAccInfo.Ptid, objAccInfo.Name)
'        GoTo EndLine
'    End If
'    SaveMICResult = True
'
'EndLine:
'    Set objAccInfo = Nothing
'End Function

''-----------------------------------------------------------------------------'
''   기능 : 미생물 배양(Culture) 결과저장
''   인수 :
''       - pIntInfo : 장비에서 수신한결과 저장 클래스
''   반환 : Ture(성공) / False(실패)
''-----------------------------------------------------------------------------'
'Public Function SaveCultureResult(ByVal pIntInfo As clsIISIntInfo) As Boolean
'    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
'    Dim strBarNo   As String    '바코드번호
'    Dim strSpcYy   As String    '검체연도
'    Dim lngSpcNo   As Long      '검체번호
'    Dim strErrMsg  As String    '에러메시지
'
'    strBarNo = pIntInfo.BarNo
'    strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
'    lngSpcNo = Val(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
'
'    Set objAccInfo = mAccInfos(strSpcYy, lngSpcNo)
'    If pIntInfo.VitekNo <> "" Then
'    If objAccInfo.CultureSaveServer(pIntInfo) = False Then
'        strErrMsg = "바코드번호: " & objAccInfo.GetBarNo & "의 서버 결과저장에 실패했습니다."
'        Error.SetLog App.EXEName, "clsIISInterface", "SaveMICResult", strErrMsg, Now
'        RaiseEvent SpcErrorX("S010", objAccInfo.GetBarNo, objAccInfo.Ptid, objAccInfo.Name)
'        GoTo EndLine
'    End If
'    End If
'    SaveCultureResult = True
'
'EndLine:
'    Set objAccInfo = Nothing
'End Function
'
'Public Function SaveComment(ByVal pIntInfo As clsIISIntInfo, ByVal pComment As String) As Boolean
'    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
'    Dim strBarNo   As String    '바코드번호
'    Dim strSpcYy   As String    '검체연도
'    Dim lngSpcNo   As Long      '검체번호
'    Dim strErrMsg  As String    '에러메시지
'
'    strBarNo = pIntInfo.BarNo
'    strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
'    lngSpcNo = Val(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
'
'    Set objAccInfo = mAccInfos(strSpcYy, lngSpcNo)
'    If objAccInfo.CommentSaveServer(pIntInfo, pComment) = False Then
''        strErrMsg = "바코드번호: " & objAccInfo.GetBarNo & "의 서버 결과저장에 실패했습니다."
''        Error.SetLog App.EXEName, "clsIISInterface", "SaveMICResult", strErrMsg, Now
''        RaiseEvent SpcErrorX("S010", objAccInfo.GetBarNo, objAccInfo.PtId, objAccInfo.Name)
'        GoTo EndLine
'    End If
'    SaveComment = True
'
'EndLine:
'    Set objAccInfo = Nothing
'
'End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호의 접수정보를 mAccInfos컬렉션에서 삭제
'   인수 :
'       - pSpcYy : 검체연도
'       - pSpcNo : 검체번호
'-----------------------------------------------------------------------------'
Public Function Remove(ByVal pSpcYy As String, ByVal pSpcNo As Long)
    If mAccInfos.Exist(pSpcYy, pSpcNo) Then
        Call mAccInfos.Remove(pSpcYy, pSpcNo)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호의 접수정보를 mAccInfos컬렉션에서 삭제
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Public Sub RemoveX(ByVal pBarNo As String)
    Dim strSpcYy As String  '검체연도
    Dim lngSpcNo As Long    '검체순번
    
    pBarNo = Format$(pBarNo, String$(SPCLEN, "#"))
    strSpcYy = Mid$(pBarNo, 1, SPCYYLEN)
    lngSpcNo = Val(Mid$(pBarNo, SPCYYLEN + 1, SPCNOLEN))
    
    Call Remove(strSpcYy, lngSpcNo)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비에서 송신한 데이터를 버퍼클래스에 저장
'   인수 :
'       - pBuffChar : 장비에서 송신한 데이터
'-----------------------------------------------------------------------------'
Public Sub AddBuffer(ByVal pBuffChar As String)
    If mBuffers.Exist(mBufCnt) Then
        mBuffers(mBufCnt).Buffers = mBuffers(mBufCnt).Buffers & pBuffChar
    Else
        Call mBuffers.Add(mBufCnt, pBuffChar)
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 버퍼클래스 Clear
'-----------------------------------------------------------------------------'
Public Sub ClearBuffer()
    mBuffers.RemoveAll
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비에서 수신한 Datalog를 저장할 파일 Open
'-----------------------------------------------------------------------------'
Private Sub SetLogFile()
    Dim strFileNm   As String   '파일명
    
    If Dir(LogPath, vbDirectory) = "" Then
        MkDir LogPath
    End If
    
    mFileNum = FreeFile
    strFileNm = LogPath & "\" & mEqpKey & "(" & Format$(Now, "YYYY-MM-DD") & ").log"
    If Dir(strFileNm) = "" Then
        Open strFileNm For Output As #mFileNum
    Else
        Open strFileNm For Append As #mFileNum
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Datalog를 파일에 저장
'   인수 :
'       - pData  : Datalog
'       - pLogFg : Datalog 구분자
'-----------------------------------------------------------------------------'
Public Sub WriteLog(ByVal pData As String, ByVal pLogFg As LogEnum)
    If pLogFg = ccPCLog Then
        Print #mFileNum, "[P C] " & pData
    Else
        Print #mFileNum, pData;
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 보관기간이 지난 로그파일, ClientDb 삭제
'-----------------------------------------------------------------------------'
Public Sub DelHistoryData()
    Dim objCAccinfo As clsIISCAccInfo   'ClientDb의 접수내역 클래스
    
    Dim strFileNm  As String     '로그파일
    Dim dtHistory  As Date       '로그파일 수정일
    Dim strErrMsg  As String     '오류메시지
    
    '## 로그파일 삭제
    strFileNm = Dir(LogPath & "\*.log", vbNormal)
    Do Until strFileNm = ""
        dtHistory = FileDateTime(LogPath & "\" & strFileNm)
        If DateDiff("d", dtHistory, Now) > mStoreDt Then
            Kill LogPath & "\" & strFileNm
        End If
        strFileNm = Dir
    Loop
    
    '## ClientDb 삭제
    Set objCAccinfo = New clsIISCAccInfo
    If objCAccinfo.DelHisClientDb(mEqpCd, mStoreDt) = False Then
        '## 오류메시지 표시
        strErrMsg = mEqpCd & " 장비의 ClientDb를 삭제중 에러가 발생했습니다."
        Error.SetLog App.EXEName, "clsIISInterface", "DelHistoryData", strErrMsg, Now
        RaiseEvent EqpError("E006")
    End If
    Set objCAccinfo = Nothing
    
    '## ClientDb 압축
    Call ClientDbCompress
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더전송시 사용되는 FrameNo를 조회
'-----------------------------------------------------------------------------'
Public Function GetFrameNo() As Long
    mFrameNo = mFrameNo + 1
    If mFrameNo = 8 Then
        mFrameNo = 0
    End If
    GetFrameNo = mFrameNo
End Function

Private Sub Class_Initialize()
    mPhase = 1
    mBufCnt = 1
    
    Set mIntNms = New clsIISIntNms
    Set mAccInfos = New clsIISAccInfos
    Set mBuffers = New clsIISBuffers
End Sub

Private Sub Class_Terminate()
    Close #mFileNum
    
    Set mBuffers = Nothing
    Set mAccInfos = Nothing
    Set mIntNms = Nothing
End Sub
