VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISIntError.cls
'   작성자  : 이상대
'   내  용  : 인터페이스중 발생하는 에러저장 클래스
'   작성일  : 2004-05-09
'   버  전  :
'       1. 1.0.1: 이상대(2004-11-18)
'          - R016 코드추가
'       2. 1.0.2: 이상대(2005-09-14)
'          - R017, R018, R019, R020 코드추가
'       3. 1.0.3: 이상대(2005-09-15)
'          - R021, R022, R023 코드추가
'-----------------------------------------------------------------------------'

Option Explicit

'## 해결방법 상수
Private Const R001 As String = "개발자에게 문의하세요."
Private Const R002 As String = "Manager -> 검사장비관련 -> 장비별 검사항목 설정 메뉴에서 해당장비에서 검사할 검사코드를 입력하세요." & vbCrLf & _
                               "이 오류를 수정하지 않으면 검사진행이 되지 않습니다."
Private Const R003 As String = "Manager -> 검사장비관련 -> 검사장비 통신설정 메뉴에서 해당장비에 대한 통신설정상태를 입력하세요."
Private Const R004 As String = "해당 검체는 처방상태 입니다. 채혈, 접수후 검사하세요."
Private Const R005 As String = "해당 검체는 채혈상태 입니다. 접수후 검사하세요."
Private Const R006 As String = "해당 검체는 결과확인, 수정된 검체입니다. 결과확인, 수정된 검체는 검사할수 없습니다."
Private Const R007 As String = "해당 검체는 접수취소된 검체입니다. 접수취소된 검체는 검사할수 없습니다."
Private Const R008 As String = "처방난 검사코드가 장비별 검사항목 설정메뉴에 등록된 검사코드인지 확인하세요."
Private Const R009 As String = "Worklist에서 해당검체를 삭제후 검사하세요."
Private Const R010 As String = "해당 장비에 설정된 포트가 다른장치에서 사용중입니다." & vbCrLf & _
                               "Manager -> 검사장비관련 -> 검사장비 통신설정 메뉴에서 포트를 변경하세요." & vbCrLf & _
                               "이 오류를 수정하지 않으면 검사진행이 되지 않습니다."
Private Const R011 As String = "프로그램 종료후 다시 실행하세요. 같은 오류가 발생하면 개발자에게 문의하세요." & vbCrLf & _
                               "이 오류를 수정하지 않으면 검사진행이 되지 않습니다."
Private Const R012 As String = "바코드번호 형식이 틀립니다. 정확히 입력해주세요."
Private Const R013 As String = "잘못된 바코드번호 입니다. 바코드번호를 확인하세요."
Private Const R014 As String = "해당 검체는 No Growth 상태입니다. Growth 등록후 검사하세요."
Private Const R015 As String = "해당 검체는 감수성 결과대상이 아닙니다. 해당 검사코드의 결과유형을 확인하세요."
Private Const R016 As String = "해당 Vitek No에 대한 접수번호가 접수내역에 존재하지 않습니다."
Private Const R017 As String = "해당 검체는 미생물 Worksheet 작성하지 않은 검체입니다. Worksheet 작성후 다시 시도하세요."
Private Const R018 As String = "설정 -> 오더파일 생성경로를 입력하세요."
Private Const R019 As String = "설정 -> 결과파일 생성경로를 입력하세요."
Private Const R020 As String = "설정 -> 오더파일명을 입력하세요."
Private Const R021 As String = "설정 -> 결과파일명 확장자를 입력하세요."
Private Const R022 As String = "설정 -> 통신설정을 입력하세요."
Private Const R023 As String = "해당 장비에 설정된 포트가 다른장치에서 사용중입니다." & vbCrLf & _
                               "설정 -> 통신설정 메뉴에서 포트를 변경하세요." & vbCrLf & _
                               "이 오류를 수정하지 않으면 검사진행이 되지 않습니다."

Private mSeq         As Long        '입력순번
Private mCode        As String      '에러코드
Private mErrDt       As String      '발생일시
Private mEqpCd       As String      '장비코드
Private mEqpNm       As String      '장비명
Private mBarNo       As String      '바코드번호
Private mPtId        As String      '환자ID
Private mName        As String      '환자명
Private mTitle       As String      '에러제목
Private mDescription As String      '에러내용
Private mResolve     As String      '해결방법

Public Property Get Seq() As Long
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As Long)
    mSeq = vData
End Property

Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal vData As String)
    mCode = vData
End Property

Public Property Get ErrDt() As String
    ErrDt = mErrDt
End Property

Public Property Let ErrDt(ByVal vData As String)
    mErrDt = vData
End Property

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get EqpNm() As String
    EqpNm = mEqpNm
End Property

Public Property Let EqpNm(ByVal vData As String)
    mEqpNm = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
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

Public Property Get Title() As String
    Title = mTitle
End Property

Public Property Let Title(ByVal vData As String)
    mTitle = vData
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal vData As String)
    mDescription = vData
End Property

Public Property Get Resolve() As String
    Resolve = mResolve
End Property

Public Property Let Resolve(ByVal vData As String)
    mResolve = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 에러코드로 에러제목을 조회
'   인수 :
'       - pCode : 에러코드
'   반혼 : 에러제목
'-----------------------------------------------------------------------------'
Public Function GetTitle(Optional ByVal pCode As String) As String
    If pCode <> "" Then mCode = pCode
    
    '## 에러코드
    '##     - E: 장비설정에 대한 에러
    '##     - S: 검체관련 에러
    '##     - X: 장비에서 수신한 에러
    If mTitle = "" Then
        Select Case mCode
            Case "E001"
                mTitle = "장비에 대한 레지스트리 정보가 없습니다."
            Case "E002"
                mTitle = "장비별 검사항목이 설정되지 않었습니다."
            Case "E003"
                mTitle = "장비에 대한 통신설정이 되어있지 않습니다."
            Case "E004"
                mTitle = "통신포트가 이미 열려있습니다."
            Case "E005"
                mTitle = "장비에 설정된 통신포트가 다른장치에서 이미 사용중입니다."
            Case "E006"
                mTitle = "ClientDb를 삭제중 에러가 발생했습니다."
            Case "E007"
                mTitle = "BEP III 장비의 오더파일 생성경로 정보가 없습니다."
            Case "E008"
                mTitle = "BEP III 장비의 경로파일 생성경로 정보가 없습니다."
            Case "E009"
                mTitle = "BEP III 장비의 오더파일명 정보가 없습니다."
            Case "E010"
                mTitle = "BEP III 장비의 결과파일명 확장자 정보가 없습니다."
            Case "E011"
                mTitle = "LASC-HSC 장비의 통신설정이 되어있지 않습니다."
            Case "E012"
                mTitle = "LASC-HSC 장비의 통신포트가 이미 열려있습니다."
            Case "E013"
                mTitle = "LASC-HSC 장비에 설정된 통신포트가 다른장치에서 이미 사용중입니다."
            Case "S001"
                mTitle = "검체에 대한 접수내역이 없습니다."
            Case "S002"
                mTitle = "검체가 접수상태가 아닙니다."
            Case "S003"
                mTitle = "검체가 접수상태가 아닙니다."
            Case "S004"
                mTitle = "검체가 접수상태가 아닙니다."
            Case "S005"
                mTitle = "검체가 접수상태가 아닙니다."
            Case "S006"
                mTitle = "검체에 대한 결과내역이 없습니다."
            Case "S007"
                mTitle = "검체에 장비로 검사할 항목이 없습니다."
            Case "S008"
                mTitle = "Worklist에 존재하는 검체입니다."
            Case "S009"
                mTitle = "ClientDb 저장에 실패했습니다."
            Case "S010"
                mTitle = "서버 결과저장에 실패했습니다."
            Case "S011"
                mTitle = "서버 결과저장에 실패했습니다."
            Case "S012"
                mTitle = "검체에 대한 QC결과내역이 없습니다."
            Case "S013"
                mTitle = "잘못된 바코드번호 입니다."
            Case "S014"
                mTitle = "Growth된 검체가 아닙니다."
            Case "S015"
                mTitle = "감수성 결과대상 검체가 아닙니다."
            Case "S016"
                mTitle = "해당 접수번호에 대해 접수내역이 없습니다."
            Case "S017"
                mTitle = "배양대상 검체가 아닙니다."
        End Select
    End If
    GetTitle = mTitle
End Function

'-----------------------------------------------------------------------------'
'   기능 : 에러에 대한 상세정보, 해결방법을 조회
'   반환 : 에러정보, 해결방법
'-----------------------------------------------------------------------------'
Public Function GetDescription() As String
    Dim strTemp As String
    
    If mTitle = "" Then Call GetTitle
    If mDescription = "" Or mResolve = "" Then
        Select Case mCode
            Case "E001"
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "E002"
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R002 & vbCrLf
            Case "E003"
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R003 & vbCrLf
            Case "E004"
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R011 & vbCrLf
            Case "E005"
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R010 & vbCrLf
            Case "E006"
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "E007"     '## BEP III 장비용
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R018 & vbCrLf
            Case "E008"     '## BEP III 장비용
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R019 & vbCrLf
            Case "E009"     '## BEP III 장비용
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R020 & vbCrLf
            Case "E010"     '## BEP III 장비용
                mDescription = "장비명: " & mEqpNm & " (" & mEqpCd & ")" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R021 & vbCrLf
            Case "E011"     '## LASC-HSC 장비용
                mDescription = "장비명: LASC-HSC" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R022 & vbCrLf
            Case "E012"     '## LASC-HSC 장비용
                mDescription = "장비명: LASC-HSC" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R011 & vbCrLf
            Case "E013"     '## LASC-HSC 장비용
                mDescription = "장비명: LASC-HSC" & vbCrLf
                mDescription = mDescription & "내  용: " & mTitle & vbCrLf
                mResolve = R023 & vbCrLf
            Case "S001"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R013 & vbCrLf
            Case "S002"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R004 & vbCrLf
            Case "S003"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R005 & vbCrLf
            Case "S004"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R006 & vbCrLf
            Case "S005"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R007 & vbCrLf
            Case "S006"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S007"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R008 & vbCrLf
            Case "S008"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R009 & vbCrLf
            Case "S009"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S010"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S011"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S012"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R001 & vbCrLf
            Case "S013"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R012 & vbCrLf
            Case "S014"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R014 & vbCrLf
            Case "S015"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R015 & vbCrLf
            Case "S016"
                mDescription = "접수번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R016 & vbCrLf
            Case "S017"
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "환 자 ID: " & mPtId & vbCrLf
                mDescription = mDescription & "환 자 명: " & mName & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
                mResolve = R017 & vbCrLf
            Case "X001"     '## Dimension AR, RXL 장비의 수신한 에러표시
                mDescription = "검체번호: " & mBarNo & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
            Case "X002"     '## AU1000 장비의 바코드리딩 에러표시
                mDescription = "Rack  No: " & mBarNo & vbCrLf
                mDescription = mDescription & "내    용: " & mTitle & vbCrLf
        End Select
    End If
    
    If InStr(mCode, "X") > 0 Then
        GetDescription = "<< 에러정보 >>" & vbCrLf & mDescription
    Else
        GetDescription = "<< 에러정보 >>" & vbCrLf & mDescription & vbCrLf & "<< 해결방법 >>" _
                            & vbCrLf & mResolve
    End If
End Function
