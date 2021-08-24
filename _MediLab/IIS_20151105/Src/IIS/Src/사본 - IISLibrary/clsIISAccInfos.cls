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
'   파일명  : clsIISAccInfos.cls (자사 LIS와 조인할때)
'   작성자  : 오세원
'   내  용  : 접수내역 컬렉션 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfos As Collection

Public Property Get Count() As Long
    Count = mAccInfos.Count
End Property

Public Property Get Item(ByVal pSpcYy As String, ByVal pSpcNo As Long) As clsIISAccInfo
    Set Item = mAccInfos(pSpcYy & DIV & CStr(pSpcNo))
End Property

Public Property Get ItemByPtID(ByVal pBarNo As String) As clsIISAccInfo
    Set ItemByPtID = mAccInfos(pBarNo)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mAccInfos.[_NewEnum]
End Property
    
'-----------------------------------------------------------------------------'
'   기능 : 접수내역 추가1 (인터페이스용)
'   인수 :
'       - pSpcYy      : 검체연도
'       - pSpcNo      : 검체번호
'       - pPtid       : 환자ID
'       - pName       : 환자명
'       - pSsn        : 주민번호
'       - pAgeDay     : 일령
'       - pSex        : 성별
'       - pDeptNm     : 진료과명
'       - pOrdDoctNm  : 처방의명
'       - pWorkarea   : Workarea
'       - pAccDt      : 접수일자
'       - pAccSeq     : 접수순번
'       - pStsCd      : 검체상태 (0:처방, 1:채혈, 2:접수, 3: 검사중, 4:중간결과, 5:완료, 6:수정, 7:접수취소)
'       - pReqTotCnt  : 총 검사항목 개수
'       - pInputCnt   : 입력된 검사항목 개수
'       - pSpcCd      : 검체코드
'       - pSpcNm      : 검체명
'       - pWardNm     : 병동명
'       - pTestDiv    : 일반, 미생물, 기타검사 구분 (0:일반검사, 1:기타검사, 2:미생물검사)
'       - pQcFg       : QC 여부 (0:일반처방, 1:QC처방)
'       - pStatFg     : 응급여부 (0:비응급, 1:응급)
'       - pBuildCd    : 건물코드
'       - pHosilId    : 호실코드
'   반환 : 접수내역 클래스
'-----------------------------------------------------------------------------'
Public Function AddInter(ByVal pSpcYy As String, ByVal pSpcNo As Long, _
                    Optional ByVal pPtID As String, Optional ByVal pName As String, _
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
        .PtId = pPtID:              .Name = pName
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
'   기능 : 접수내역 추가1 (인터페이스용)
'   인수 :
'       - pAccInfo : 접수내역 객체
'-----------------------------------------------------------------------------'
Public Sub AddInterX(ByVal pAccInfo As clsIISAccInfo)
    mAccInfos.Add pAccInfo, pAccInfo.SpcYy & DIV & CStr(pAccInfo.SpcNo)
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 접수내역 추가1 (인터페이스용)
'   인수 :
'       - pAccInfo : 접수내역 객체
'-----------------------------------------------------------------------------'
Public Sub AddInterXByPtID(ByVal pAccInfo As clsIISAccInfo)
    mAccInfos.Add pAccInfo, pAccInfo.PtId
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'   인수 :
'       - pSpcYy : 검체연도
'       - pSpcNo : 검체번호
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pSpcYy As String, ByVal pSpcNo As Long)
    mAccInfos.Remove pSpcYy & DIV & CStr(pSpcNo)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    For i = mAccInfos.Count To 1 Step -1
        mAccInfos.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   인수 :
'       - pSpcYy : 검체연도
'       - pSpcNo : 검체번호
'   반환 : True(존재) / Flase(비존재)
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

'-----------------------------------------------------------------------------'
'   기능 : 해당 Index의 존재유무를 반환
'   인수 :
'       - pSpcYy : 검체연도
'       - pSpcNo : 검체번호
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function ExistByPtID(ByVal pPtID As String) As Boolean
    Dim objAccInfo As clsIISAccInfo
    
    For Each objAccInfo In mAccInfos
        With objAccInfo
            If pPtID = .PtId Then
                ExistByPtID = True
                GoTo EndLine
            End If
        End With
    Next
    ExistByPtID = False

EndLine:
    Set objAccInfo = Nothing
End Function

Private Sub Class_Initialize()
    Set mAccInfos = New Collection
End Sub

Private Sub Class_Terminate()
    Set mAccInfos = Nothing
End Sub

