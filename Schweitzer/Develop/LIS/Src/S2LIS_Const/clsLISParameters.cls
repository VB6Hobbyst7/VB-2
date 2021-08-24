VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISParameters"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public P_RealTestMicSpecial           As Boolean          '관련검사(미생물/특수검사포함여부)
Public P_ReasonCdFG                 As Boolean          '수정시 결과코드 입력
Public P_ApplyCalculation           As Boolean
Public P_UseBarcodeSystem           As Boolean
Public P_MICSelectedByUser          As Boolean
Public P_UseLabCommentSystem        As Boolean
Public P_BBSResultReview            As Boolean
Public P_BBSCollection              As Boolean
Public P_IncludeBBSSystem           As Boolean
Public P_ElectronicSignature        As Boolean          '기타검사 전자싸인여부
Public P_MicElectronicSign          As Boolean          '미생물 전자싸인여부
Public P_AllResultReview            As Boolean          '전체결과 조회 여부
Public P_Arlet_PanicChk               As Boolean          '성별/나이별 Panic/Arlet체크
Public p_UseWSBatchRst         As Boolean          'Worksheet일괄등록사용여부
Public p_UseInstrBatchRst             As Boolean          '장비별 일괄등록 사용여부
Public P_DiffFg                     As Boolean          'Diff 체크여부
Public P_ItemSeqFG                  As Boolean          '통계출력 순서를 이용한 통계작성
Public P_RealPrinter                As Boolean          '즉시프린트사용여부
Public P_PrinterChkFg               As Boolean          '즉시프린트사용시 rptfg업데이트 여부
Public P_ApplyBuildingInfo          As Boolean
Public P_ErBarcodeAutoPrint         As Boolean
Public P_UseReceptForSearch         As Boolean
Public P_IncludeAPSSystem           As Boolean
Public P_ImageSystem                As Boolean          '이미지 시스템 사용 여부
Public P_CheckSugaCode              As Boolean
Public P_IncludePOCSystem           As Boolean
Public P_IncludeCMTSystem           As Boolean
Public P_ApplyMicMidVerify          As Boolean          '미생물 중간결과 적용 여부
Public P_MicSensiReport             As Boolean          '미생물 MIC Sensi 출력여부
Public P_ABOCHK                     As Boolean          '결과등록시 ABO체크 사용여부
Public P_ESRResultChk                 As Boolean          'ESR 검사결과 체크
Public P_MornCollection               As Boolean          '아침채혈여부
Public P_PayDtUsed                  As Boolean          '외래의 수납개념 사용
Public P_CollectMultiSpc            As String
Public P_ResultReviewFlag           As String           '0-임상/해부/혈액, 1-임상, 2-해부, 3-임상/해부, _
                                                         4-혈액, 5-임상/혈액, 6-해부/혈액
Public P_PatientIdFormat            As String
Public P_HOSPITALNAME               As String
Public P_ReviewStartDate            As String           '전체결과 조회 시작일
Public P_MicSensiTestCd             As String           '미생물 MIC Sensi 검사코드
Public P_ErDeptCd                   As String           '응급실 부서코드
Public P_DiffSpcCd                  As String           'Diff Count 검체코드
Public P_AFBSENSCD                  As String           '결핵균검사코드
Public P_AFPCULTCD                  As String           '결핵배양검사코드
Public P_ABOTestCD                  As String           '혈액형코드
Public P_RHTestCD                   As String           'Rh검사코드
Public P_NoResultReport             As String
Public P_NoWorkareaReport           As String           '결과지 찍지 않는 workarea
Public P_HOSPITALADDR               As String
Public P_HOSPITALADDR_ENG           As String
Public P_SLIDE_SERVER_PATH            As String
Public P_SLIDE_DB_PATH                As String  '"C:\Schweitzer\Image\"
Public P_SLIDE_CLIENT_PATH            As String '= "C:\Data\SlideImage\"
Public P_ESRCode                      As String           'ESR 검사코드
Public P_HctCode                      As String           '헤마토크리트 검사코드
Public P_POCCode                      As String           'POC검사항목코드
Public P_PtidLength                 As String           '환자 길이
Public P_SpcYyLength                As Long
Public P_SpcNoLength                As Long
Public P_VerifyListInterval         As Long             '결과보고 대기 리스트 검색기간
Public P_BLOCK_SIZE                   As Long
'Public AttributeBBS As Boolean

'## By M.G.Choi (2006-09-06)
'   - 결과지에 출력되는 확인자명
Public P_HOSPITALCHIP               As String           '결과지에 출력되는 확인자명

'## By M.G.Choi (2006-09-06)
'   - 결과지에 출력되는 바닥글
Public P_HOSPITALCOMMENT            As String           '결과지 바닥글 Comment

'## By M.G.Choi (2007-07-04)
Public W_HeadMessage                As String
Public W_XPOS                       As Long
Public W_YPOS                       As Long
Public W_BackStyle                  As String
Public W_ShowDelay                  As Long
Public W_MessageHeight              As Long
Public W_MessageWidth               As Long
Public W_HFontBold                  As Boolean
Public W_DFontBold                  As Boolean
Public W_DFontSize                  As Long


Private Sub Class_Initialize()

    If Dir(INIPath) = "" Then
        MsgBox INIPath & " 파일유무를 확인하세요.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    W_HeadMessage = GetConstVariable("W_HeadMessage")
    W_XPOS = GetConstVariable("W_XPOS")
    W_YPOS = GetConstVariable("W_YPOS")
    W_BackStyle = GetConstVariable("W_BackStyle")
    W_ShowDelay = GetConstVariable("W_ShowDelay")
    W_MessageHeight = GetConstVariable("W_MessageHeight")
    W_MessageWidth = GetConstVariable("W_MessageWidth")
    W_HFontBold = GetConstVariable("W_HFontBold")
    W_DFontBold = GetConstVariable("W_DFontBold")
    W_DFontSize = GetConstVariable("W_DFontSize")
    
     'TRUE/FALSE
    P_RealTestMicSpecial = GetConstVariable("P_RealTestMicSpecial")       ' True
    P_ReasonCdFG = GetConstVariable("P_ReasonCdFG")                   ' True
    P_ApplyCalculation = GetConstVariable("P_ApplyCalculation")       ' True
    P_UseBarcodeSystem = GetConstVariable("P_UseBarcodeSystem")       ' True
    P_MICSelectedByUser = GetConstVariable("P_MICSelectedByUser")     ' True
    P_UseLabCommentSystem = GetConstVariable("P_UseLabCommentSystem") ' True
    P_BBSResultReview = GetConstVariable("P_BBSResultReview")         ' True
    P_BBSCollection = GetConstVariable("P_BBSCollection")             ' True
    P_IncludeBBSSystem = GetConstVariable("P_IncludeBBSSystem")       ' True
    P_ElectronicSignature = GetConstVariable("P_ElectronicSignature") ' True
    P_MicElectronicSign = GetConstVariable("P_MicElectronicSign")     ' True
    P_AllResultReview = GetConstVariable("P_AllResultReview")         ' True
    P_Arlet_PanicChk = GetConstVariable("P_Arlet_PanicChk")               ' True
    p_UseWSBatchRst = GetConstVariable("p_UseWSBatchRst")   ' True
    p_UseInstrBatchRst = GetConstVariable("p_UseInstrBatchRst")           ' True
    
    P_DiffFg = GetConstVariable("P_DiffFg")                           ' False
    P_ItemSeqFG = GetConstVariable("P_ItemSeqFG")                     ' False
    P_RealPrinter = GetConstVariable("P_RealPrinter")                 ' False
    P_PrinterChkFg = GetConstVariable("P_PrinterChkFg")               ' False
    P_ApplyBuildingInfo = GetConstVariable("P_ApplyBuildingInfo")     ' False
    P_ErBarcodeAutoPrint = GetConstVariable("P_ErBarcodeAutoPrint")   ' False
    P_UseReceptForSearch = GetConstVariable("P_UseReceptForSearch")   ' False
    P_IncludeAPSSystem = GetConstVariable("P_IncludeAPSSystem")       ' False
    P_ImageSystem = GetConstVariable("P_ImageSystem")                 ' False
    P_CheckSugaCode = GetConstVariable("P_CheckSugaCode")             ' False
    P_IncludePOCSystem = GetConstVariable("P_IncludePOCSystem")       ' False
    P_IncludeCMTSystem = GetConstVariable("P_IncludeCMTSystem")       ' False
    P_ApplyMicMidVerify = GetConstVariable("P_ApplyMicMidVerify")     ' False
    P_MicSensiReport = GetConstVariable("P_MicSensiReport")           ' False
    P_ABOCHK = GetConstVariable("P_ABOCHK")                           ' False
    P_ESRResultChk = GetConstVariable("P_EsrResultChk")                   ' False
    P_MornCollection = GetConstVariable("P_MornCollection")               ' False
    P_PayDtUsed = GetConstVariable("P_PayDtUsed")                     ' True
    
    'String형
    P_CollectMultiSpc = GetConstVariable("P_CollectMultiSpc")         ' "0"
    P_ResultReviewFlag = GetConstVariable("P_ResultReviewFlag")       ' "0"
    P_PatientIdFormat = GetConstVariable("P_PatientIdFormat")         ' "00000000"
    P_HOSPITALNAME = GetConstVariable("P_HOSPITALNAME")               ' "가톨릭대학교 성모자애병원"
    P_ReviewStartDate = GetConstVariable("P_ReviewStartDate")         ' ""
    P_MicSensiTestCd = GetConstVariable("P_MicSensiTestCd")           ' "L4062"
    P_ErDeptCd = GetConstVariable("P_ErDeptCd")                       ' "'ER'"
    P_DiffSpcCd = GetConstVariable("P_DiffSpcCd")                     ' "1B"
    P_AFBSENSCD = GetConstVariable("P_AFBSENSCD")                     ' ""
    P_AFPCULTCD = GetConstVariable("P_AFPCULTCD")                     ' ""
    P_ABOTestCD = GetConstVariable("P_ABOTestCD")                     ' ""
    P_RHTestCD = GetConstVariable("P_RHTestCD")                       ' ""
    P_NoResultReport = GetConstVariable("P_NoResultReport")           ' ""
    P_NoWorkareaReport = GetConstVariable("P_NoWorkareaReport")       ' ""
    P_HOSPITALADDR = GetConstVariable("P_HOSPITALADDR")               ' "인천시 부평구성모자애병원"
    
    '** 일단 Fix ㅡ.ㅡ 낭중에 INI File
    P_HOSPITALCHIP = "김영숙" 'GetConstVariable_New("P_HOSPITALCHIP")
    P_HOSPITALCOMMENT = "* 본 검사실은 대한진단검사의학회(KSLIM)의 신임인증을 받은 우수검사실로서 결과의 정확성 및 신뢰도를 보증합니다."
    P_HOSPITALADDR_ENG = "P.O. Box 77, Presbyterian Medical Center Chonju, Cholla Pukdo, Korea (Tel. 063-230-8114)"
    
        
    P_SLIDE_SERVER_PATH = GetConstVariable("P_SLIDE_SERVER_PATH")         ' "C:\Schweitzer\Server\"
    P_SLIDE_DB_PATH = GetConstVariable("P_SLIDE_DB_PATH")                 ' "C:\Schweitzer\Image\"
    P_SLIDE_CLIENT_PATH = GetConstVariable("P_SLIDE_CLIENT_PATH")         ' "C:\Data\SlideImage\"
    P_ESRCode = GetConstVariable("P_ESRCode")                             ' "L1030"
    P_HctCode = GetConstVariable("P_HctCode")                             ' "L1020"
    P_POCCode = GetConstVariable("P_POCCode")                             ' "L3710"
    P_PtidLength = GetConstVariable("P_PtidLength")
    'Long형
    P_SpcYyLength = GetConstVariable("P_SpcYyLength")                 ' 2
    P_SpcNoLength = GetConstVariable("P_SpcNoLength")                 ' 9
    P_VerifyListInterval = GetConstVariable("P_VerifyListInterval")   ' -1
                      ' 8
    P_BLOCK_SIZE = GetConstVariable("P_BLOCK_SIZE")                       ' 16384
'    AttributeBBS = GetConstVariable("AttributeBBS")
    
'    'TRUE/FALSE
'    P_RealTestMicSpecial = True          '관련검사:특수검사/미생물포함여부
'    P_ReasonCdFG = True
'    P_ApplyCalculation = True
'    P_UseBarcodeSystem = True
'    P_MICSelectedByUser = True
'    P_UseLabCommentSystem = True
'    P_BBSResultReview = True
'    P_BBSCollection = True
'    P_IncludeBBSSystem = True
'    P_ElectronicSignature = True
'    P_MicElectronicSign = True
'    P_AllResultReview = True
'    P_Arlet_PanicChk = True
'    p_UseWSBatchRst = True
'    p_UseInstrBatchRst = True
'
'    P_DiffFg = False
'    P_ItemSeqFG = False
'    P_RealPrinter = False
'    P_PrinterChkFg = False
'    P_ApplyBuildingInfo = False
'    P_ErBarcodeAutoPrint = False
'    P_UseReceptForSearch = False
'    P_IncludeAPSSystem = False
'    P_ImageSystem = False
'    P_CheckSugaCode = False
'    P_IncludePOCSystem = False
'    P_IncludeCMTSystem = False
'    P_ApplyMicMidVerify = False
'    P_MicSensiReport = False
'    P_ABOCHK = False
'    P_ERSResultChk = False
'    P_MornCollection = False
'
'    'String형
'    P_CollectMultiSpc = "0"
'    P_ResultReviewFlag = "0"
'    P_PatientIdFormat = "00000000"
'    P_HOSPITALNAME = "가톨릭대학교 성모자애병원"
'    P_ReviewStartDate = ""
'    P_MicSensiTestCd = "L4062"
'    P_ErDeptCd = "'ER'"
'    P_DiffSpcCd = "1B"
'    P_AFBSENSCD = ""
'    P_AFPCULTCD = ""
'    P_ABOTestCD = ""
'    P_RHTestCD = ""
'    P_NoResultReport = ""
'    P_NoWorkareaReport = ""
'    P_HOSPITALADDR = "인천시 부평구성모자애병원"
'    P_SLIDE_SERVER_PATH = "C:\Schweitzer\Server\"
'    P_SLIDE_DB_PATH = "C:\Schweitzer\Image\"
'    P_SLIDE_CLIENT_PATH = "C:\Data\SlideImage\"
'    P_ESRCode = "L1030"
'    P_HctCode = "L1020"
'    P_POCCode = "L3710"
'
'    'Long형
'    P_SpcYyLength = 2
'    P_SpcNoLength = 9
'    P_VerifyListInterval = -1
'    P_PtidLength = 8
'    P_BLOCK_SIZE = 16384
End Sub

Private Function GetConstVariable(ByVal qKeyString As String) As Variant
    Dim strTmp  As String
    
    If Dir(INIPath) = "" Then Exit Function
    strTmp = medGetINI("LIS_CONST", qKeyString, INIPath)
    
    Select Case medGetP(strTmp, 2, LINE_DIV)
        Case "0": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
        Case "1": GetConstVariable = CLng(medGetP(strTmp, 1, LINE_DIV))
        Case "2": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
    End Select
    
End Function

'Private Function GetWarnVariable(ByVal qKeyString As String) As Variant
'    Dim strTmp  As String
'
'    If Dir(INIPath) = "" Then Exit Function
'    strTmp = medGetINI("WARN_MSG", qKeyString, INIPath)
'
'    Select Case medGetP(strTmp, 2, LINE_DIV)
'        Case "0": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
'        Case "1": GetConstVariable = CLng(medGetP(strTmp, 1, LINE_DIV))
'        Case "2": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
'    End Select
'
'End Function

