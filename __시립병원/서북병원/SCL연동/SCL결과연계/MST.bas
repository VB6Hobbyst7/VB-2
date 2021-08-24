Attribute VB_Name = "Mst"
Option Explicit
    
Type LabRtnMstRec
    LabParCod   As String   '母 코드(Parent Code)   (K-1)
    LabSeqNum   As String   'SeqNo                  (K-2)
    LabMbrCod   As String   '子 코드(Member Code)   (D-1)
    LabSonCod   As String   '孫 코드
End Type
 
'------------------------------------------------------
'37) 검사기기 등록 정보   LabMchMst
'------------------------------------------------------
Type LabMchMstRec
    LabMchCod       As String       ' Key
    LabMchNam       As String       ' 명칭
    LabMchShtNam    As String       ' 약어
    LabMchUidCod    As String       ' 담당자
    LabMchDev       As String       ' COM port
    LabMchBps       As String       ' Baud rate
    LabMchPar       As String       ' Parity
    LabMchDat       As String       ' Data bit
    LabMchSta       As String       ' Start bit
    LabMchSto       As String       ' Stop bit
    LabMchRmk       As String       ' 비고
    LabTstCod       As String       '기기의 검사종류
    LabTryYon       As String
End Type
    
'------------------------------------------------------
'38) 검사의뢰처 등록 정보   LabSclMst
'------------------------------------------------------
Type LabSclMstRec
    LabSclCod       As String       ' Key
    LabSclNam       As String       ' 명칭
    LabSclShtNam    As String       ' 약어
    LabSclRmk       As String       ' 비고
End Type
    
'------------------------------------------------------
'38) 검체 등록 정보   LabSpmMst
'------------------------------------------------------
Type LabSpmMstRec
    LabSpmCod       As String       ' Key
    LabSpmSeq       As String       ' 순서
    LabSpmNam       As String       ' 명칭
    LabSpmShtNam    As String       ' 약어
    LabSpmRmk       As String       ' 비고
End Type
    
'------------------------------------------------------
'38) 검사종류 등록 정보   LabTstMst
'------------------------------------------------------
Type LabTstMstRec
    LabTstCod       As String       ' Key
    LabTstNam       As String       ' 명칭
    LabTstShtNam    As String       ' 약어
    LabTstRmk       As String       ' 비고
End Type
    
'------------------------------------------------------
'34) 검사정보 NewLabMst
'------------------------------------------------------
Type NewLabMstRec
    LabCod     As String        '검사코드 (K-1)
    LabCodNam  As String        '1명칭       (D-1)
    LabShtNam   As String       '2약자 명칭
    LabSpmCod  As String        '3검체코드
    LabComMax  As String        '4공통상한치
    LabComLow  As String        '5공통하한치 (D-5)
    LabComRef   As String       '6공통표준값
    LabMalMax  As String        '7남성상한치
    LabMalLow  As String        '8남성하한치
    LabMalRef   As String       '9남성표준값
    LabFmlMax  As String        '0여성상한치 (D-10)
    LabFmlLow  As String        '1여성하한치
    LabFmlRef   As String       '2여성표준값
    LabMzhUnt  As String        '3결과단위
    LabSclCod  As String        '4수탁여부
    LabRltTyp   As String       '5검사결과유형 (D-15) S:single-line, M:multi-line, C:combo-box
    LabDefRlt   As String       '6Default 검사 결과
    LabRltOpt   As String       '7선택가능한 검사결과 들 ("-|±|1+|2+|3+|4+")
    LabMaxLen   As String       '8검사결과의 최대 길이
    LabMaxLin   As String       '9검사결과의 최대 줄수
    LabRltSeq   As String       '0검사 결과 입력 화면에서 순서 (D-20)
    LabJbsSeq   As String       '1검사 접수시 화면에서의 순서
    LabUidCod   As String       '2검사 담당자 코드
    LabMchNum   As String       '3검사 장비 번호
    LabMchCod   As String       '4검사 장비에서의 코드
    LabSlpTyp1  As String       '5검사대분류 (D-25)
    LabSlpTyp2  As String       '6검사소분류
    LabPrtYon   As String       '7워크리스트나, 처방전에 찍을 항목인지
    LabMulJbs   As String       '8수량이 여러번 order 나왔을때 한꺼번에 접수할건지 아니면 한번씩 접수할건지?
    LabViwYon   As String       '9화면 Display 여부
    LabSotTyp   As String       '0검사종목별 통계분류
    LabAdpDte   As String       '적용일
    LabExpDte   As String       '종료일
    LabJngGbn   As String       '1정도관리 방법
    LabTrmVal   As String       '2유효기간
    LabDltMax   As String       '3Delta 상한
    LabDltLow   As String       '4Delta 하한
    LabDlyDay   As String       '5평균 결과조회일수
    LabPanMax   As String       'Panic 상한
    LabPanLow   As String       'Panic 하한
    
End Type
    
    
'------------------------------------------------------
'34) 검사기계에서 넘어오는 검사 코드와
'    우리가 사용하는 검사코드의 맷칭
'------------------------------------------------------
Type MchCodMstRec
    MchCod      As String       ' Key
    MchTstCod   As String       ' Key
    MchLabCod   As String       ' 1     우리코드
    MchPos      As String       ' 2     프로그램에서 보여주는 순서
End Type
    
'----------------------------------------------------
'인터페이스할 검사 장비를 선택한 후 이를 저장
'----------------------------------------------------
Type LabMchSelRec
    LabMchKey As String
    LabMchCod As String
End Type
    
'정도관리 임시파일
Type DelPanRec
    ResChtNum As String     'K-1   Chart No.
    ResLabCod As String     'K-2   LabCode
    ResPatNam As String     'D-1   Patient Name
    ResSexVal As String     'D-2   Sex
    ResAgeVal As String     'D-3   Age
    ResDelVal As String     'D-4   Delta value
    ResPanVal As String     'D-5   Panic value
    ResOldDte As String     'D-6   전 검사일
    ResDayVal As String     'D-7   현재일 value
    ResOldVal As String     'D-8   전일 value
    ResSpmCod As String     'D-9   동일검체여부
End Type
'///////////////////////////////////////////////////////////
'/// 종진 출력시 엑셀(보고서) 위치 설정
'//////////////////////////////////////////////////////////
Type MexExcMstRec
    MexMexTyp As String        'key
    MexPagNum As String        'key
    MexLabCod As String        'key
    MexFleNam As String
    MexExcPos As String
    MexHozCnt As String
    MexVrtCnt As String
    MexAcpCod As String
End Type
    
    
'------------------------------------------------------
'New Blood Bank 에서 혈액 마스터로 사용함
'------------------------------------------------------
Type BldMstRec
    BldCodNum   As String   'K-1    혈액번호
    BldPakNam   As String   'D-1    혈액제제
    BldTypNam   As String   'D-2    혈액형
    BldRh       As String   'D-3    Rh
    BldInDtm    As String   'D-4    혈액입고일
    BldAboDtm   As String   'D-5    혈액폐기일
    BldStaFlg   As String   'D-6    혈액상태
    BldAcpDte   As String   'D-7    접수일자
    BldAcpNum   As String   'D-8    접수번호
    BldEndDtm   As String   'D-9    완료일
    BldAboCmt   As String   'D-10   폐기사유
End Type
    
    
'------------------------------------------------------
'New Blood Bank 에서 혈액불출정보로 사용함
'------------------------------------------------------
Type BldInfRec
    BldAcpDte   As String   'K-1    접수일자
    BldAcpCod   As String   'K-2    접수코드
    BldAcpNum   As String   'K-3    접수번호
    BldOcmNum   As String   'D-1    내원번호
    BldChtNum   As String   'D-2    챠트번호
    BldNum      As String   'D-3    혈액번호
    BldPatTyp   As String   'D-4    환자혈액형
    BldPatRh    As String   'D-5    환자혈액Rh
    BldManMat   As String   'D-6    주교차시험
    BldSubMat   As String   'D-7    부교차시험
    BldMatUid   As String   'D-8    교차시험자
    BldSndUid   As String   'D-9    혈액불출자
    BldRevUid   As String   'D-10   혈액인수자
    BldSndDtm   As String   'D-11   혈액불출일
End Type
    
Public Sub BldInfLoad(sPrmValue As String, tPrmData As BldInfRec)

    On Error GoTo BldInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.BldAcpDte = vVal(i)
    i = i + 1
    tPrmData.BldAcpCod = vVal(i)
    i = i + 1
    tPrmData.BldAcpNum = vVal(i)
    i = i + 1
    tPrmData.BldOcmNum = vVal(i)
    i = i + 1
    tPrmData.BldChtNum = vVal(i)
    i = i + 1
    tPrmData.BldNum = vVal(i)
    i = i + 1
    tPrmData.BldPatTyp = vVal(i)
    i = i + 1
    tPrmData.BldPatRh = vVal(i)
    i = i + 1
    tPrmData.BldManMat = vVal(i)
    i = i + 1
    tPrmData.BldSubMat = vVal(i)
    i = i + 1
    tPrmData.BldMatUid = vVal(i)
    i = i + 1
    tPrmData.BldSndUid = vVal(i)
    i = i + 1
    tPrmData.BldRevUid = vVal(i)
    i = i + 1
    tPrmData.BldSndDtm = vVal(i)
    
    
    Exit Sub

BldInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub BldInfStore(sPrmKey As String, sPrmValue As String, tPrmData As BldInfRec)

    
    sPrmKey = tPrmData.BldAcpDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.BldAcpCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.BldAcpNum & Chr(5)
    
    sPrmValue = tPrmData.BldOcmNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldChtNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldPatRh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldManMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldSubMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldMatUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldSndUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldRevUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldSndDtm & Chr(5)
End Sub

    
Public Sub BldMstLoad(sPrmValue As String, tPrmData As BldMstRec)

    On Error GoTo BldMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.BldCodNum = vVal(i)
    i = i + 1
    tPrmData.BldPakNam = vVal(i)
    i = i + 1
    tPrmData.BldTypNam = vVal(i)
    i = i + 1
    tPrmData.BldRh = vVal(i)
    i = i + 1
    tPrmData.BldInDtm = vVal(i)
    i = i + 1
    tPrmData.BldAboDtm = vVal(i)
    i = i + 1
    tPrmData.BldStaFlg = vVal(i)
    i = i + 1
    tPrmData.BldAcpDte = vVal(i)
    i = i + 1
    tPrmData.BldAcpNum = vVal(i)
    i = i + 1
    tPrmData.BldEndDtm = vVal(i)
    i = i + 1
    tPrmData.BldAboCmt = vVal(i)
    
    
    Exit Sub

BldMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub BldMstStore(sPrmKey As String, sPrmValue As String, tPrmData As BldMstRec)

    
    sPrmKey = tPrmData.BldCodNum & Chr(5)
    
    sPrmValue = tPrmData.BldPakNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldTypNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldRh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldInDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldAboDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldStaFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldAcpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldAcpNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldEndDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.BldAboCmt & Chr(5)
    
End Sub

    
Public Sub DelPanLoad(sPrmValue As String, tPrmData As DelPanRec)

    On Error GoTo DelPanLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.ResChtNum = vVal(i)
    i = i + 1
    tPrmData.ResLabCod = vVal(i)
    i = i + 1
    tPrmData.ResPatNam = vVal(i)
    i = i + 1
    tPrmData.ResSexVal = vVal(i)
    i = i + 1
    tPrmData.ResAgeVal = vVal(i)
    i = i + 1
    tPrmData.ResDelVal = vVal(i)
    i = i + 1
    tPrmData.ResPanVal = vVal(i)
    i = i + 1
    tPrmData.ResOldDte = vVal(i)
    i = i + 1
    tPrmData.ResDayVal = vVal(i)
    i = i + 1
    tPrmData.ResOldVal = vVal(i)
    i = i + 1
    tPrmData.ResSpmCod = vVal(i)
    
    
    Exit Sub

DelPanLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DelPanStore(sPrmKey As String, sPrmValue As String, tPrmData As DelPanRec)

    
    sPrmKey = tPrmData.ResChtNum & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ResLabCod & Chr(5)
    
    sPrmValue = tPrmData.ResPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ResSexVal & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ResAgeVal & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ResDelVal & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ResPanVal & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ResOldDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ResDayVal & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ResOldVal & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ResSpmCod & Chr(5)
    
    
End Sub

    
Public Sub LabMchMstLoad(sPrmValue As String, tPrmData As LabMchMstRec)

    On Error GoTo LabMchMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LabMchCod = vVal(i)
    i = i + 1
    tPrmData.LabMchNam = vVal(i)
    i = i + 1
    tPrmData.LabMchShtNam = vVal(i)
    i = i + 1
    tPrmData.LabMchUidCod = vVal(i)
    i = i + 1
    tPrmData.LabMchDev = vVal(i)
    i = i + 1
    tPrmData.LabMchBps = vVal(i)
    i = i + 1
    tPrmData.LabMchPar = vVal(i)
    i = i + 1
    tPrmData.LabMchDat = vVal(i)
    i = i + 1
    tPrmData.LabMchSta = vVal(i)
    i = i + 1
    tPrmData.LabMchSto = vVal(i)
    i = i + 1
    tPrmData.LabMchRmk = vVal(i)
    i = i + 1
    tPrmData.LabTstCod = vVal(i)
    i = i + 1
    tPrmData.LabTryYon = vVal(i)
    Exit Sub

LabMchMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LabMchMstStore(sPrmKey As String, sPrmValue As String, tPrmData As LabMchMstRec)

    
    sPrmKey = tPrmData.LabMchCod & Chr(5)
    
    sPrmValue = tPrmData.LabMchNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabMchShtNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabMchUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabMchDev & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabMchBps & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabMchPar & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabMchDat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabMchSta & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabMchSto & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabMchRmk & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabTstCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabTryYon & Chr(5)
    
End Sub

    
Public Sub LabMchSelLoad(sPrmValue As String, tPrmData As LabMchSelRec)

    On Error GoTo LabMchSelLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.LabMchKey = vVal(i)
    i = i + 1
    tPrmData.LabMchCod = vVal(i)
    
    Exit Sub

LabMchSelLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LabMchSelStore(sPrmKey As String, sPrmValue As String, tPrmData As LabMchSelRec)

    
    sPrmKey = tPrmData.LabMchKey & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LabMchCod & Chr(5)
    
    sPrmValue = "" & Chr(5)
    
End Sub

    
Public Sub LabRtnMstLoad(sPrmValue As String, tPrmData As LabRtnMstRec)

    On Error GoTo LabRtnMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LabParCod = vVal(i)
    i = i + 1
    tPrmData.LabSeqNum = vVal(i)
    i = i + 1
    tPrmData.LabMbrCod = vVal(i)
    i = i + 1
    tPrmData.LabSonCod = vVal(i)
    Exit Sub

LabRtnMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LabRtnMstStore(sPrmKey As String, sPrmValue As String, tPrmData As LabRtnMstRec)

    sPrmKey = tPrmData.LabParCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LabSeqNum & Chr(5)
    
    sPrmValue = tPrmData.LabMbrCod & Chr(5) & tPrmData.LabSonCod & Chr(5)
End Sub

    
Public Sub LabSclMstLoad(sPrmValue As String, tPrmData As LabSclMstRec)

    On Error GoTo LabSclMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LabSclCod = vVal(i)
    i = i + 1
    tPrmData.LabSclNam = vVal(i)
    i = i + 1
    tPrmData.LabSclShtNam = vVal(i)
    i = i + 1
    tPrmData.LabSclRmk = vVal(i)
    Exit Sub

LabSclMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LabSclMstStore(sPrmKey As String, sPrmValue As String, tPrmData As LabSclMstRec)

    sPrmKey = tPrmData.LabSclCod & Chr(5)
    
    sPrmValue = tPrmData.LabSclNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabSclShtNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabSclRmk & Chr(5)
End Sub

    
Public Sub LabSpmMstLoad(sPrmValue As String, tPrmData As LabSpmMstRec)

    On Error GoTo LabSpmMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LabSpmCod = vVal(i)
    i = i + 1
    tPrmData.LabSpmSeq = vVal(i)
    i = i + 1
    tPrmData.LabSpmNam = vVal(i)
    i = i + 1
    tPrmData.LabSpmShtNam = vVal(i)
    i = i + 1
    tPrmData.LabSpmRmk = vVal(i)
    Exit Sub

LabSpmMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LabSpmMstStore(sPrmKey As String, sPrmValue As String, tPrmData As LabSpmMstRec)

    sPrmKey = tPrmData.LabSpmCod & Chr(5)
    
    sPrmValue = tPrmData.LabSpmSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabSpmNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabSpmShtNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabSpmRmk & Chr(5)
End Sub

    
Public Sub LabTstMstLoad(sPrmValue As String, tPrmData As LabTstMstRec)

    On Error GoTo LabTstMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LabTstCod = vVal(i)
    i = i + 1
    tPrmData.LabTstNam = vVal(i)
    i = i + 1
    tPrmData.LabTstShtNam = vVal(i)
    i = i + 1
    tPrmData.LabTstRmk = vVal(i)
    Exit Sub

LabTstMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LabTstMstStore(sPrmKey As String, sPrmValue As String, tPrmData As LabTstMstRec)

    sPrmKey = tPrmData.LabTstCod & Chr(5)
    
    sPrmValue = tPrmData.LabTstNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabTstShtNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LabTstRmk & Chr(5)
End Sub

    
Public Sub MchCodMstLoad(sPrmValue As String, tPrmData As MchCodMstRec)

    On Error GoTo MchCodMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.MchCod = vVal(i)
    i = i + 1
    tPrmData.MchTstCod = vVal(i)
    i = i + 1
    tPrmData.MchLabCod = vVal(i)
    i = i + 1
    tPrmData.MchPos = vVal(i)
    Exit Sub

MchCodMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MchCodMstStore(sPrmKey As String, sPrmValue As String, tPrmData As MchCodMstRec)

    sPrmKey = tPrmData.MchCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MchTstCod & Chr(5)
    
    sPrmValue = tPrmData.MchLabCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MchPos & Chr(5)
End Sub

    
Public Sub MexExcMstLoad(sPrmValue As String, tPrmData As MexExcMstRec)

    On Error GoTo MexExcMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    ''If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.MexMexTyp = vVal(i)
    i = i + 1
    tPrmData.MexPagNum = vVal(i)
    i = i + 1
    tPrmData.MexLabCod = vVal(i)
    i = i + 1
    tPrmData.MexFleNam = vVal(i)
    i = i + 1
    tPrmData.MexExcPos = vVal(i)
    i = i + 1
    tPrmData.MexHozCnt = vVal(i)
    i = i + 1
    tPrmData.MexVrtCnt = vVal(i)
    i = i + 1
    tPrmData.MexAcpCod = vVal(i)
    
    
    Exit Sub

MexExcMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MexExcMstStore(sPrmKey As String, sPrmValue As String, tPrmData As MexExcMstRec)

    
    sPrmKey = tPrmData.MexMexTyp & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MexPagNum & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MexLabCod & Chr(5)
    
    sPrmValue = tPrmData.MexFleNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MexExcPos & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MexHozCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MexVrtCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MexAcpCod & Chr(5)
    
End Sub

    
Public Sub NewLabMstLoad(sPrmValue As String, tPrmLabData As NewLabMstRec)

    On Error GoTo NewLabMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmLabData.LabCod = vVal(i)
    i = i + 1
    tPrmLabData.LabCodNam = vVal(i)
    i = i + 1
    tPrmLabData.LabShtNam = vVal(i)
    i = i + 1
    tPrmLabData.LabSpmCod = vVal(i)
    i = i + 1
    tPrmLabData.LabComMax = vVal(i)
    i = i + 1
    tPrmLabData.LabComLow = vVal(i)
    i = i + 1
    tPrmLabData.LabComRef = vVal(i)
    i = i + 1
    tPrmLabData.LabMalMax = vVal(i)
    i = i + 1
    tPrmLabData.LabMalLow = vVal(i)
    i = i + 1
    tPrmLabData.LabMalRef = vVal(i)
    i = i + 1
    tPrmLabData.LabFmlMax = vVal(i)
    i = i + 1
    tPrmLabData.LabFmlLow = vVal(i)
    i = i + 1
    tPrmLabData.LabFmlRef = vVal(i)
    i = i + 1
    tPrmLabData.LabMzhUnt = vVal(i)
    i = i + 1
    tPrmLabData.LabSclCod = vVal(i)
    i = i + 1
    tPrmLabData.LabRltTyp = vVal(i)
    i = i + 1
    tPrmLabData.LabDefRlt = vVal(i)
    i = i + 1
    tPrmLabData.LabRltOpt = vVal(i)
    i = i + 1
    tPrmLabData.LabMaxLen = vVal(i)
    i = i + 1
    tPrmLabData.LabMaxLin = vVal(i)
    i = i + 1
    tPrmLabData.LabRltSeq = vVal(i)
    i = i + 1
    tPrmLabData.LabJbsSeq = vVal(i)
    i = i + 1
    tPrmLabData.LabUidCod = vVal(i)
    i = i + 1
    tPrmLabData.LabMchNum = vVal(i)
    i = i + 1
    tPrmLabData.LabMchCod = vVal(i)
    i = i + 1
    tPrmLabData.LabSlpTyp1 = vVal(i)
    i = i + 1
    tPrmLabData.LabSlpTyp2 = vVal(i)
    i = i + 1
    tPrmLabData.LabPrtYon = vVal(i)
    i = i + 1
    tPrmLabData.LabMulJbs = vVal(i)
    i = i + 1
    tPrmLabData.LabViwYon = vVal(i)
    i = i + 1
    tPrmLabData.LabSotTyp = vVal(i)
    i = i + 1
    tPrmLabData.LabAdpDte = vVal(i)
    i = i + 1
    tPrmLabData.LabExpDte = vVal(i)
    i = i + 1
    tPrmLabData.LabJngGbn = vVal(i)
    i = i + 1
    tPrmLabData.LabTrmVal = vVal(i)
    i = i + 1
    tPrmLabData.LabDltMax = vVal(i)
    i = i + 1
    tPrmLabData.LabDltLow = vVal(i)
    i = i + 1
    tPrmLabData.LabDlyDay = vVal(i)
    i = i + 1
    tPrmLabData.LabPanMax = vVal(i)
    i = i + 1
    tPrmLabData.LabPanLow = vVal(i)
    
    Exit Sub

NewLabMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub NewLabMstStore(sPrmKey As String, sPrmValue As String, tPrmLabData As NewLabMstRec)

    
    sPrmKey = tPrmLabData.LabCod & Chr(5)
    
    sPrmValue = tPrmLabData.LabCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabShtNam & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSpmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabComMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabComLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabComRef & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMalMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMalLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMalRef & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabFmlMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabFmlLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabFmlRef & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMzhUnt & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSclCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabRltTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabDefRlt & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabRltOpt & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMaxLen & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMaxLin & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabRltSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabJbsSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMchNum & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMchCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSlpTyp1 & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSlpTyp2 & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabPrtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMulJbs & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabViwYon & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSotTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabJngGbn & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabTrmVal & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabDltMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabDltLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabDlyDay & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabPanMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabPanLow & Chr(5)
    
End Sub

    

Public Sub NewLabMstRead(sPrmLabCod As String, tPrmNewLabData As NewLabMstRec, Optional sPrmOdrDte As String = "99999999")

    Dim scurKey As String
    Dim sCmpKey As String
    Dim sRetval As String
    
    Dim NewLabData As NewLabMstRec
    
    sCmpKey = sPrmLabCod & Chr(5)
    scurKey = sCmpKey
    scurKey = mSetPrev("NewLabMst", scurKey)
    Do
        scurKey = mReadPrev("NewLabMst", scurKey, sCmpKey, sRetval)
        If scurKey = "" Then Exit Do
        
        Call NewLabMstLoad(sRetval, NewLabData)
    
        If sPrmOdrDte >= CDouble(NewLabData.LabAdpDte) And sPrmOdrDte <= CDouble(NewLabData.LabExpDte) Then
            tPrmNewLabData = NewLabData
        End If
    Loop
    
End Sub

Public Sub LabSpmMstRead(sPrmSpmCod As String, tPrmLabSpmData As LabSpmMstRec)

    Dim scurKey As String
    Dim sCmpKey As String
    Dim sRetval As String
    
    sCmpKey = sPrmSpmCod & Chr(5)
    scurKey = sCmpKey
    scurKey = mSetReadEqual("LabSpmMst", scurKey, sRetval)
    Call LabSpmMstLoad(sRetval, tPrmLabSpmData)
    
End Sub

Public Sub LabTstMstRead(sPrmTstCod As String, tPrmLabTstData As LabTstMstRec)

    Dim scurKey As String
    Dim sCmpKey As String
    Dim sRetval As String
    
    sCmpKey = sPrmTstCod & Chr(5)
    scurKey = sCmpKey
    scurKey = mSetReadEqual("LabTstMst", scurKey, sRetval)
    Call LabTstMstLoad(sRetval, tPrmLabTstData)
    
End Sub

