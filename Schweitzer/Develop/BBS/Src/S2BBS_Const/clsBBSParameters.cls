VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSParameters"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================
'      병원 정보관련 정의
'================================
Public HOSPITAL_MAIN        As String               '병원명
Public HOSPITAL_NAME        As String               '혈액원 명칭
Public HOSPITAL_TEL1        As String               '전화번호1
Public HOSPITAL_TEL2        As String               '전화번호2
Public HOSPITAL_ADDR        As String               '주소
Public HOSPITAL_ADDR2       As String               '주소2
Public HOSPITAL_FAX         As String               '팩스
Public BLOOD_LABEL_CNT      As String
Public BBS_PTID_LENGTH      As String               '환자ID 길이(성모자애:8)
Public BLOOD_DEPTCD         As String               '혈액은행 Deptcd

'================================
'      병원별 프로그램 사용여부
'================================
Public BLOOD_DEL_USED       As Boolean              '출고전표 사용여부
Public DonorUserFg          As Boolean              '헌혈업무 사용여부
Public TRANS_REQUIRE_USED   As Boolean              '수혈요청 사용여부
Public TransReactionUsed    As Boolean              '수혈부작용등록여부
Public ABO_DoubleChk        As Boolean              '혈액형더블 체크여부
Public BloodSplitUsed       As Boolean              '혈액분획사용여부
Public BloodTransfer        As Boolean              '혈액이동사용여부
Public BloodLocalDelivery   As Boolean              '혈액Local 출고사용여부
Public DonationPaper        As Boolean              '헌혈증관리
Public AttributeBBS         As Boolean              '처방속성의 사용(BBS테이블)
Public TEST_FOR_PHERSIS     As Boolean              '혈액은행 Deptcd

Public BLOOD_STORE_BARCODE_USED As Boolean               '혈액입고시 혈액원코드 사용(바코드로만 입력)
Public USE_DONOR_INFORM As Boolean '헌혈 약식기능 사용(헌혈자등록, 혈액등록, 헌혈자조회, 지정취소기능만 제공)

Private Sub Class_Initialize()
    
    If Dir(INIPath) = "" Then
        MsgBox INIPath & " 파일유무를 확인하세요.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    '전역상수
    AttributeBBS = GetConstVariable("AttributeBBS")             ' True
    DonorUserFg = GetConstVariable("DonorUserFg")               ' True
    TEST_FOR_PHERSIS = GetConstVariable("TEST_FOR_PHERSIS")     ' True
    BLOOD_DEL_USED = GetConstVariable("BLOOD_DEL_USED")         ' False
    TRANS_REQUIRE_USED = GetConstVariable("TRANS_REQUIRE_USED") ' False
    TransReactionUsed = GetConstVariable("TransReactionUsed")   ' False
    ABO_DoubleChk = GetConstVariable("ABO_DoubleChk")           ' False
    BloodSplitUsed = GetConstVariable("BloodSplitUsed")         ' False
    BloodTransfer = GetConstVariable("BloodTransfer")           ' False
    BloodLocalDelivery = GetConstVariable("BloodLocalDelivery") ' False
    DonationPaper = GetConstVariable("DonationPaper")           ' False
    BLOOD_STORE_BARCODE_USED = GetConstVariable("BLOOD_STORE_BARCODE_USED")
    USE_DONOR_INFORM = GetConstVariable("USE_DONOR_INFORM")
    'String
    HOSPITAL_MAIN = GetConstVariable("HOSPITAL_MAIN")           ' "동국대학교병원"
    HOSPITAL_NAME = GetConstVariable("HOSPITAL_NAME")           ' "동국대학교병원 혈액은행"
    HOSPITAL_TEL1 = GetConstVariable("HOSPITAL_TEL1")           ' "            "
    HOSPITAL_TEL2 = GetConstVariable("HOSPITAL_TEL2")           ' "            "
    HOSPITAL_ADDR = GetConstVariable("HOSPITAL_ADDR")           ' "경주시 동국대학교 병원 혈액은행"
    HOSPITAL_ADDR2 = GetConstVariable("HOSPITAL_ADDR2")         ' "경주시 동국대학교 병원 혈액은행"
    HOSPITAL_FAX = GetConstVariable("HOSPITAL_FAX")             ' "             "
    BLOOD_DEPTCD = GetConstVariable("BLOOD_DEPTCD")             ' "CP"
    BBS_PTID_LENGTH = GetConstVariable("BBS_PTID_LENGTH")       ' "8"
    'Long
    BLOOD_LABEL_CNT = GetConstVariable("BLOOD_LABEL_CNT")       ' 1
End Sub

Private Function GetConstVariable(ByVal qKeyString As String) As Variant
    Dim strTmp  As String
    
    If Dir(INIPath) = "" Then Exit Function
    strTmp = medGetINI("BBS_CONST", qKeyString, INIPath)
    
    Select Case medGetP(strTmp, 2, LINE_DIV)
        Case "0": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
        Case "1": GetConstVariable = CLng(medGetP(strTmp, 1, LINE_DIV))
        Case "2": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
    End Select
    
End Function



