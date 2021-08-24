Attribute VB_Name = "modEasyCare"
Option Explicit

Public gOnline_Test As String
Public gServerPath As String
Public gIFUser As String
Public gIFName As String

Public giIndex  As Long

'암센터
Public Const gXml_S01 = "PG_SRL.SLP91_S01"
Public Const gXml_S02 = "PG_SRL.SLP91_S02"
Public Const gXml_S03 = "PG_SRL.SLP91_S03"
Public Const gXml_S04 = "PG_SRL.SLP91_S04"
Public Const gXml_S07 = "PG_SRL.SLP91_S07"
Public Const gXml_S10 = "PG_SRL.SLP91_S10"
Public Const gXml_S18 = "PG_SRL.SLP91_S18"
Public Const gXml_S24 = "PG_SRL.SLP91_S24"
Public Const gXml_U07 = "PG_SRL.SLP91_U01"

'이대목동
'-- 로그인
Public Const gXml_LOGIN = "PKG_MSE_LM_INTERFACE.PC_MSE_USER_SELECT"
'-- 바코드 조회
'Public Const gXml_ORDER_SELECT = "PKG_MSE_LM_INTERFACE.PC_MSE_ORDER_SELECT"
'-- 결과저장
'Public Const gXml_RESULT_UPLOAD = "PKG_MSE_LM_INTERFACE.PC_MSE_INTERFACE_SAVE"


'서울대병원
Public Const gXml_ORDER_SELECT = "PKG_LAB.INTERFACE_S29"
Public Const gXml_BAR_SELECT = "PKG_LAB.INTERFACE_S29"      '바코드
'Public Const gXml_BAR_SELECT = "PKG_LAB.INTERFACE_S32"      '바코드
Public Const gXml_DAY_SELECT = "PKG_LAB.INTERFACE_S28"      'ONE Day
Public Const gXml_TERM_SELECT = "PKG_LAB.INTERFACE_S23"     'From~To Date


'서울대
Public Const gXml_RESULT_UPLOAD = "PKG_LAB.INTERFACE_I02"



Public Const gXml_ACPT_SPNO = "PKG_MSE_LM_INTERFACE.PC_MSE_INS_ACPT_SPNO"

Public gOrderExam As String


Type Exam_Select
    TST_CD      As String
    TST_CNT     As Integer
End Type

Public gExam_Select()   As Exam_Select

Type PatInfo_Select
'    TST_CD      As String
'    TST_NM      As String
'    TST_FRCT_CD As String
'    ACPTNO_1    As String
'    PT_NO       As String
'    PT_NM       As String
'    Sex         As String
'    SPC_CD_1    As String
'    ORD_SITE    As String
'    TST_CLS     As String
'    RERUN       As String
'    TST_FRCT_CD1    As String
'    HSP_CLS     As String
'    ACPT_DTETM  As String
'    Age         As String
'    OK          As Integer

    SPCM_NO                 As String   '검체번호
    ACPT_DTM                As String   '접수일자
    EXM_ACPT_NO             As String   '접수번호
    PT_NO                   As String   '환자번호
    PT_NM                   As String   '환자이름
    SEX_TP_CD               As String   '성별
    PT_BRDY_DT              As String   '생년월일
    PT_HME_DEPT_CD          As String   '환자진료과
    WD_DEPT_CD              As String   '병동
    EXM_CD                  As String   '검사코드
    TH1_SPCM_CD             As String   '검체코드
    HR24_URN_EXM_TM         As String   '24시간소변검사시간
    HR24_URN_EXM_VLM_CNTE   As String   '24시간소변검사부피내용
    RPRN_EXM_CD             As String   '판넬코드
    EXM_PRGR_STS_CD         As String   '결과상태값
    OK                      As Integer
End Type

Public gPatInfo_Select  As PatInfo_Select


Type PatInfo_Term_Select
    ACPT_DTE            As String  ' 검사일자
    ACPTNO_1            As String  ' 접수번호
    PT_NO               As String  ' 환자번호
    PATNAME             As String  ' 환자명
    SEX                 As String  ' 성별
    AGE                 As String  ' 나이
    SPC_NO              As String  ' 검체번호
    ORD_SITE            As String  ' 발행처
    WRK_UNT_CD          As String  ' 작업코드
    TST_CD              As String  ' 검사코드
    TST_SNM             As String  ' 검사명
    SPC_NM              As String  ' 검체명
    TST_STAT            As String  ' 상태\
    ACPT_DTETM          As String  ' 접수일시
    DOCTORNOTE_YN       As String  ' DOCTORNOTE_YN
    SND_ARVL_NO_CNTE    As String  ' SND_ARVL_NO_CNTE
    DEXM_YENU           As String  ' 연번호
    OK                  As Integer
End Type

Public gPInfo_T_Sel  As PatInfo_Term_Select


'======================== gnuh_emr ======================================

'Get_QCList 바코드번호, 구분 - QC 정보 불러오기
'Get_QCWorkList 검사일자, 장비번호 - QC WorkList 불러오기
'Online_QCResult "99910084349", "C061", "664887", "20091008151515", 5, " L63011  L63012  L63013  L6371   L6377   ", " 1.1 2.2 3.3 4.4 5.1 "
' - QC 결과 전송

Type Order_Select
    SPC_NO      As String
    PT_NO       As String
    PT_NM       As String
    ACPT_DTE    As String
    ACPT_NO     As String
    TST_CD      As String
    WRK_UNT     As String
    TST_DTE     As String
    TST_STAT    As String
    WD_NO       As String
    SPC_NM      As String
    SPC_CD      As String
    OK          As Integer
End Type
Public gOrder_Select As Order_Select
Public gOrder_List() As Order_Select
Public gWork_Select() As Order_Select
'Public giIndex  As Long

Type Patient_Info
    PTNO        As String
    PATNAME     As String
    SEX         As String
    AGE         As String
    DPCD        As String
    WD_NO       As String
    SPC_CD      As String
    SPC_NM      As String
    ACPT_NO     As String
    ACPT_DTM    As String
    TST_STAT    As String
    OK          As Integer
End Type
Public gPatient_Info As Patient_Info

'추가 변수 start==========================================
Type QC_Info
    INST_DTM    As String
    LOT_NO      As String
    TST_CD      As String
    EQUIP_CD    As String
    CTRL_CD     As String
    LOT_NO1     As String
    BARCODE_CD  As String
    USE_YN      As String
    OK          As Integer
End Type

Public gOnline_Ret As String
'======================== gnuh_emr ======================================


Public Sub Clear_XML_Exam()
    giIndex = -1
    ReDim gExam_Select(0)
End Sub

Public Sub Clear_XML_PInfo()
'    gPatInfo_Select.ACPT_DTETM = ""
'    gPatInfo_Select.ACPTNO_1 = ""
'    gPatInfo_Select.Age = ""
'    gPatInfo_Select.HSP_CLS = ""
'    gPatInfo_Select.OK = -1
'    gPatInfo_Select.ORD_SITE = ""
'    gPatInfo_Select.PT_NM = ""
'    gPatInfo_Select.PT_NO = ""
'    gPatInfo_Select.RERUN = ""
'    gPatInfo_Select.Sex = ""
'    gPatInfo_Select.SPC_CD_1 = ""
'    gPatInfo_Select.TST_CD = ""
'    gPatInfo_Select.TST_CLS = ""
'    gPatInfo_Select.TST_FRCT_CD = ""
'    gPatInfo_Select.TST_FRCT_CD1 = ""
'    gPatInfo_Select.TST_NM = ""
    
    gPatInfo_Select.SPCM_NO = ""                       '검체번호
    gPatInfo_Select.ACPT_DTM = ""                      '접수일자
    gPatInfo_Select.EXM_ACPT_NO = ""                   '접수번호
    gPatInfo_Select.PT_NO = ""                         '환자번호
    gPatInfo_Select.PT_NM = ""                         '환자이름
    gPatInfo_Select.SEX_TP_CD = ""                     '    성별
    gPatInfo_Select.PT_BRDY_DT = ""                    '생년월일
    gPatInfo_Select.PT_HME_DEPT_CD = ""                '진료과
    gPatInfo_Select.WD_DEPT_CD = ""                    '병동
    gPatInfo_Select.EXM_CD = ""                        '검사코드
    gPatInfo_Select.OK = 0
    gPatInfo_Select.TH1_SPCM_CD = ""                   '검체코드
    gPatInfo_Select.HR24_URN_EXM_TM = ""               '24Hour 소뱐검사시간
    gPatInfo_Select.HR24_URN_EXM_VLM_CNTE = ""         '24Hour 소뱐검사부피내용
    gPatInfo_Select.RPRN_EXM_CD = ""                   '판넬코드
    gPatInfo_Select.EXM_PRGR_STS_CD = ""               '결과상태값
    
End Sub

Public Function Online_TLA(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String

    Dim xDoc        As MSXML2.DOMDocument
    Dim xPE         As MSXML2.IXMLDOMParseError
    Dim sRetStr     As String
    Dim sFileName   As String
    Dim sParam      As String
    Dim strErrText  As String
    
    Online_TLA = ""
    sFileName = "Res"
    
    sParam = TLA_Param(asProc, asDate1, asDate2, asBarcode)
    
    sRetStr = Online_XML_Qry(asProc, sParam)
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    Set xDoc = New MSXML2.DOMDocument
    If xDoc.Load(App.PATH & "\XML\" & sFileName & ".xml") Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_S04
'            Clear_XML_TLA
            display_online_parsing_TLAInfo xDoc.childNodes, 0
        Case gXml_S24
'            Clear_XML_TLA
            gIFName = ""
            display_online_parsing_User xDoc.childNodes, 0
        End Select
'        Display_Online_Parsing_Test xDoc.childNodes, 0
    Else
        ' 문서를 로드하지 못했습니다.
       ' ParseError 개체를 가져옵니다
        Set xPE = xDoc.parseError
        With xPE
        
            strErrText = "Your XML Document failed to load" & _
                         "due the following error." & vbCrLf & _
                         "Error #: " & .errorCode & ": " & xPE.reason & _
                         "Line #: " & .Line & vbCrLf & _
                         "Line Position: " & .linepos & vbCrLf & _
                         "Position In File: " & .filepos & vbCrLf & _
                         "Source Text: " & .srcText & vbCrLf & _
                         "Document URL: " & .URL
        End With

'        SaveXML_Data strErrText
    End If
    
    Xml_Log gOnline_Test, "TLA"
    
    
    Set xPE = Nothing
    Set xDoc = Nothing
    
    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_TLA = LEFT(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If
    
End Function

Public Function Online_XML(ByVal asProc As String, ByVal asSpcno As String, Optional ByVal asDIV As String, Optional ByVal PID As String, Optional ByVal PWD As String, _
                           Optional ByVal FDate As String, Optional ByVal TDate As String, Optional ByVal TestS As String) As String
                           
    Dim xmlDoc      As MSXML2.DOMDocument30
    Dim xPE         As MSXML2.IXMLDOMParseError
    Dim strErrText  As String
    Dim sRetStr     As String
    Dim sFileName   As String
    Dim sParam      As String
    Dim Nodes       As Object
    
    Dim nodeBook As IXMLDOMElement
    Dim nodeId As IXMLDOMAttribute
    Dim xNode As MSXML2.IXMLDOMNode
    Dim namedNodeMap As IXMLDOMNamedNodeMap
    Dim Child_Node As MSXML2.IXMLDOMNodeList
    Dim i As Integer
    Dim J As Integer
    
    
    Online_XML = ""
    sFileName = "Res"
    
    '파라미터를 만든다
    sParam = Select_Param(asProc, asSpcno, PID, PWD, FDate, TDate, TestS)
    
    'MsgBox sParam
    
    '파라미터를 서버에 전송항여 리턴값을 받아온다.
    sRetStr = Online_XML_Qry(asDIV, sParam)
    
    '==================================================================================
    
    Dim strXmlName  As String
    Dim varRcvData  As Variant
    
'        ReDim Preserve XmlWork.ACPT_DTE(asCnt)
'        ReDim Preserve XmlWork.ACPTNO_1(asCnt)
'        ReDim Preserve XmlWork.PT_NO(asCnt)
'        ReDim Preserve XmlWork.PATNAME(asCnt)
'        ReDim Preserve XmlWork.SEX(asCnt)
'        ReDim Preserve XmlWork.AGE(asCnt)
'        ReDim Preserve XmlWork.SPC_NO(asCnt)
'        ReDim Preserve XmlWork.ORD_SITE(asCnt)
'        ReDim Preserve XmlWork.WRK_UNT_CD(asCnt)
'        ReDim Preserve XmlWork.TST_CD(asCnt)
'        ReDim Preserve XmlWork.TST_SNM(asCnt)
'        ReDim Preserve XmlWork.SPC_NM(asCnt)
'        ReDim Preserve XmlWork.TST_STAT(asCnt)
'        ReDim Preserve XmlWork.ACPT_DTETM(asCnt)
'        ReDim Preserve XmlWork.DOCTORNOTE_YN(asCnt)
'        ReDim Preserve XmlWork.SND_ARVL_NO_CNTE(asCnt)
'        ReDim Preserve XmlWork.DEXM_YENU(asCnt)
            
            
'    sRetStr = ""
'    sRetStr = sRetStr & "<NewDataSet>" & vbCrLf
'    sRetStr = sRetStr & "  <Table>" & vbCrLf
'    sRetStr = sRetStr & "    <SPC_NO>19060530634</SPC_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <ACPT_DTE>2019-06-05 11:28:30</ACPT_DTE>" & vbCrLf
'    sRetStr = sRetStr & "    <ACPTNO_1>8104</ACPTNO_1>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_NO>10890698</PT_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <PATNAME>전경숙</PATNAME>" & vbCrLf
'    sRetStr = sRetStr & "    <SEX>F</SEX>" & vbCrLf
'    sRetStr = sRetStr & "    <AGE>1958-10-10</AGE>" & vbCrLf
'    sRetStr = sRetStr & "    <ORD_SITE>IMP</ORD_SITE>" & vbCrLf
'    sRetStr = sRetStr & "    <WD_DEPT_CD>062</WD_DEPT_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <WRK_UNT_CD>L251321</WRK_UNT_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <TH1_SPCM_CD>1BA4</TH1_SPCM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <HR24_URN_EXM_TM>24</HR24_URN_EXM_TM>" & vbCrLf
'    sRetStr = sRetStr & "    <HR24_URN_EXM_VLM_CNTE>0</HR24_URN_EXM_VLM_CNTE>" & vbCrLf
'    sRetStr = sRetStr & "    <TST_CD>L251311</TST_CD>" & vbCrLf
'    sRetStr = sRetStr & "   <EXM_PRGR_STS_CD>C</EXM_PRGR_STS_CD>" & vbCrLf
'    sRetStr = sRetStr & " </Table>" & vbCrLf
'    sRetStr = sRetStr & "  <Table>" & vbCrLf
'    sRetStr = sRetStr & "    <SPCM_NO>19060530634</SPCM_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <ACPT_DTM>2019-06-05 11:28:30</ACPT_DTM>" & vbCrLf
'    sRetStr = sRetStr & "    <EXM_ACPT_NO>8104</EXM_ACPT_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_NO>10890698</PT_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_NM>전경숙</PT_NM>" & vbCrLf
'    sRetStr = sRetStr & "    <SEX_TP_CD>F</SEX_TP_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_BRDY_DT>1958-10-10</PT_BRDY_DT>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_HME_DEPT_CD>IMP</PT_HME_DEPT_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <WD_DEPT_CD>062</WD_DEPT_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <EXM_CD>L251322</EXM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <TH1_SPCM_CD>1BA4</TH1_SPCM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <HR24_URN_EXM_TM>24</HR24_URN_EXM_TM>" & vbCrLf
'    sRetStr = sRetStr & "    <HR24_URN_EXM_VLM_CNTE>0</HR24_URN_EXM_VLM_CNTE>" & vbCrLf
'    sRetStr = sRetStr & "    <RPRN_EXM_CD>L251311</RPRN_EXM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <EXM_PRGR_STS_CD>C</EXM_PRGR_STS_CD>" & vbCrLf
'    sRetStr = sRetStr & "  </Table>" & vbCrLf
'    sRetStr = sRetStr & "  <Table>" & vbCrLf
'    sRetStr = sRetStr & "    <SPCM_NO>19060530634</SPCM_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <ACPT_DTM>2019-06-05 11:28:30</ACPT_DTM>" & vbCrLf
'    sRetStr = sRetStr & "    <EXM_ACPT_NO>8104</EXM_ACPT_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_NO>10890698</PT_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_NM>전경숙</PT_NM>" & vbCrLf
'    sRetStr = sRetStr & "    <SEX_TP_CD>F</SEX_TP_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_BRDY_DT>1958-10-10</PT_BRDY_DT>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_HME_DEPT_CD>IMP</PT_HME_DEPT_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <WD_DEPT_CD>062</WD_DEPT_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <EXM_CD>L251325</EXM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <TH1_SPCM_CD>1BA4</TH1_SPCM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <HR24_URN_EXM_TM>24</HR24_URN_EXM_TM>" & vbCrLf
'    sRetStr = sRetStr & "    <HR24_URN_EXM_VLM_CNTE>0</HR24_URN_EXM_VLM_CNTE>" & vbCrLf
'    sRetStr = sRetStr & "    <RPRN_EXM_CD>L251311</RPRN_EXM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <EXM_PRGR_STS_CD>C</EXM_PRGR_STS_CD>" & vbCrLf
'    sRetStr = sRetStr & "  </Table>" & vbCrLf
'    sRetStr = sRetStr & "  <Table>" & vbCrLf
'    sRetStr = sRetStr & "    <SPCM_NO>19060530634</SPCM_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <ACPT_DTM>2019-06-05 11:28:30</ACPT_DTM>" & vbCrLf
'    sRetStr = sRetStr & "    <EXM_ACPT_NO>8104</EXM_ACPT_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_NO>10890698</PT_NO>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_NM>전경숙</PT_NM>" & vbCrLf
'    sRetStr = sRetStr & "    <SEX_TP_CD>F</SEX_TP_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_BRDY_DT>1958-10-10</PT_BRDY_DT>" & vbCrLf
'    sRetStr = sRetStr & "    <PT_HME_DEPT_CD>IMP</PT_HME_DEPT_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <WD_DEPT_CD>062</WD_DEPT_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <EXM_CD>L251326</EXM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <TH1_SPCM_CD>1BA4</TH1_SPCM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <HR24_URN_EXM_TM>24</HR24_URN_EXM_TM>" & vbCrLf
'    sRetStr = sRetStr & "    <HR24_URN_EXM_VLM_CNTE>0</HR24_URN_EXM_VLM_CNTE>" & vbCrLf
'    sRetStr = sRetStr & "    <RPRN_EXM_CD>L251311</RPRN_EXM_CD>" & vbCrLf
'    sRetStr = sRetStr & "    <EXM_PRGR_STS_CD>C</EXM_PRGR_STS_CD>" & vbCrLf
'    sRetStr = sRetStr & "  </Table>" & vbCrLf
'    sRetStr = sRetStr & "</NewDataSet>" & vbCrLf

'sRetStr = ""
'
'    sRetStr = sRetStr & "<?xml version='1.0' encoding='utf-8'?>"
'    sRetStr = sRetStr & "<NewDataSet>"
'    sRetStr = sRetStr & "    <Table0>"
'    sRetStr = sRetStr & "        <PT_NO><![CDATA[19400255]]></PT_NO>"
'        sRetStr = sRetStr & "    <PATNAME><![CDATA[김수지]]></PATNAME>"
'        sRetStr = sRetStr & "    <TST_CD><![CDATA[L25617]]></TST_CD>"
'        sRetStr = sRetStr & "    <AN><![CDATA[341]]></AN>"
'        sRetStr = sRetStr & "    <DAY_YN><![CDATA[N]]></DAY_YN>"
'        sRetStr = sRetStr & "    <ORD_SITE><![CDATA[IMI]]></ORD_SITE>"
'            sRetStr = sRetStr & "<TST_DTE><![CDATA[2020-06-03 오후 3:55:44]]></TST_DTE>"
'            sRetStr = sRetStr & "<TST_STAT><![CDATA[1]]></TST_STAT>"
'            sRetStr = sRetStr & "<BIRTHDAY><![CDATA[1985-08-24]]></BIRTHDAY>"
'            sRetStr = sRetStr & "<SEX><![CDATA[F]]></SEX>"
'            sRetStr = sRetStr & "<RCNT_RSLT><![CDATA[]]></RCNT_RSLT>"
'        sRetStr = sRetStr & "</Table0>"
'    sRetStr = sRetStr & "</NewDataSet>"
    
    
    Call XmlSelect_Free

    If InStr(1, sRetStr, "<NewDataSet>") > 0 Then
        varRcvData = Split(sRetStr, "<Table")
    End If
    
    'MsgBox UBound(varRcvData)
    
    strXmlName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyymmdd") & ".xml"
    
    'MsgBox sRetStr
    'MsgBox strXmlName
    
    Call SetXMLData(strXmlName, sRetStr)
    
    'Call DisplayNode_InfoS(App.PATH & "\XML\" & strXmlName, UBound(varRcvData))
    'Call DisplayNode_InfoS("D:\프로젝트\VB\__NU\인천성모병원_NU\참고\worklist.xml", 10)
    
    '==================================================================================
    
    
    'Call Xml_Log(sRetStr, sFileName)
    
    If sRetStr = "" Then
        Exit Function
    End If
    
    Dim strData As String
    
    Erase strRecvData
    intBufCnt = 1000
    ReDim Preserve strRecvData(1)
    J = 1
    For i = 1 To Len(sRetStr)
        strData = Mid(sRetStr, i, 1)
        Select Case strData
            Case vbCr
            Case vbLf
                J = J + 1
                ReDim Preserve strRecvData(J)
            Case Else
                strRecvData(J) = strRecvData(J) & strData
        End Select
    Next
    
    Set xmlDoc = New MSXML2.DOMDocument30
    If UBound(strRecvData) >= 1 Then
        ' Data Load, Start Parsing
        Select Case asProc
            Case gXml_S03
                Clear_XML_PInfo
                display_online_parsing_PatInfo xmlDoc.childNodes, 0
            Case gXml_S07
                Clear_XML_Exam
                display_online_parsing_ExamCode xmlDoc.childNodes, 0
            
            Case gXml_LOGIN
                'Clear_XML_PInfo
                'display_online_parsing_Login xmlDoc.childNodes, 0
                
                Online_XML = mGetP(mGetP(strRecvData(4), 2, "<"), 2, ">")
            Case gXml_ORDER_SELECT
                Clear_XML_PInfo
                'display_online_parsing_PatInfo xmlDoc.childNodes, 0
                'display_online_PatInfo xmlDoc.childNodes, 0 '
                'Online_XML = strRecvData(3)
            
                Call DisplayNode_PatInfo(App.PATH & "\XML\" & strXmlName, UBound(varRcvData))
                
                Online_XML = XmlInfo.PT_NO  'strRecvData(0)
            
            Case gXml_BAR_SELECT 'gXml_BAR_SELECT
                Clear_XML_PInfo
                'display_online_PatInfo xmlDoc.childNodes, 0
                
                Call DisplayNode_PatInfo(App.PATH & "\XML\" & strXmlName, UBound(varRcvData))
                
                Online_XML = XmlInfo.PT_NO  'strRecvData(0)
                            
            Case gXml_TERM_SELECT
                Clear_XML_PInfo
                'Call display_online_PatInfo_Term(xmlDoc.childNodes, 0)
                Call DisplayNode_Worklist(App.PATH & "\XML\" & strXmlName, UBound(varRcvData))

                Online_XML = XmlWork.SPC_NO(0)
            
        End Select
    Else
        ' 문서를 로드하지 못했습니다.
       ' ParseError 개체를 가져옵니다
        Set xPE = xmlDoc.parseError
        With xPE
        
            strErrText = "Your XML Document failed to load" & _
                         "due the following error." & vbCrLf & _
                         "Error #: " & .errorCode & ": " & xPE.reason & _
                         "Line #: " & .Line & vbCrLf & _
                         "Line Position: " & .linepos & vbCrLf & _
                         "Position In File: " & .filepos & vbCrLf & _
                         "Source Text: " & .srcText & vbCrLf & _
                         "Document URL: " & .URL
            
            Debug.Print strErrText
        
        End With

        SaveXML_Data strErrText
    End If
    
    
    Set xPE = Nothing
    Set xmlDoc = Nothing
    
    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_XML = LEFT(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If
    
    Kill App.PATH & "\XML\" & strXmlName

End Function

'high-level interface
Public Function Online_XML_Qry(ByVal asStrDiv As String, ByVal asParam As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    'Call soapclient.mssoapinit("DocSample1.wsdl", "TestService1", "TestServicePort")
    
    
    oSOAP.MSSoapInit gHOSP.APIURL & "?wsdl"
        
    strDiv = asStrDiv
    sParam = asParam
    
    SaveXML_Data "[Use Proc => " & strDiv & " ]" & sParam
    
    'Call soapclient.AddNumbers(2,3) '웹서비스에 정의된 메소드 호출
    'send = oSOAP.LMService(strDiv, sParam)
    '예제) this.WcfService.ServiceReturnCustomType(sType, "", "XML", "N", strXML);
    
    send = oSOAP.ServiceReturnCustomType("GETQUERY", "", "XML", "N", sParam)
    
    
    
    SaveXML_Data "[Return Proc => " & strDiv & " ]" & send
    Online_XML_Qry = send
    Set oSOAP = Nothing
    DoEvents
    Exit Function
    
ErrHandle:
    
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
End Function

'low-level interface
'Public Function Online_XML_Qry_Low(ByVal asStrDiv As String, ByVal asParam As String) As String
'    Dim Serializer  As SoapSerializer30 '전송할 데이터를 SOAP XML형태로
'    Dim Reader      As SoapReader30 '받은 데이터를 XML 형태로
'    Dim strMethod   As String
'
'    'strMethod = LMService(strDiv, sParam)
'
'    Set Connector = New HttpConnector30 '해당 주소로 연결
'
'    Connector.Property("EndPointURL") = gHOSP.APIURL
'    Connector.Connect
'    Connector.Property("SoapAction") = "uri:" & Method '
'    Connector.BeginMessage
'
'    Set Serializer = New SoapSerializer30
'
'    Serializer.Init Connector.InputStream
'    MsgBox ("SOAP 통신 데이터생성")
'    Serializer.StartEnvelope
'    Serializer.StartBody
'    Serializer.StartElement "getRecommendation", CALC_NS, , "nstemp"
'    Serializer.StartElement "data"
'    Serializer.WriteString Text1.Text
'    Serializer.EndElement
'    Serializer.EndElement
'    Serializer.EndBody
'    Serializer.EndEnvelope
'    Connector.EndMessage
'    On Error Resume Next
'    MsgBox ("SOAP 통신 결과 출력")
'    Set Reader = New SoapReader30
'    Reader.Load Connector.OutputStream
'    richText.Text = Reader.Body.XML
'    MsgBox Reader.Body.XML
'End Function


Private Function Select_Param(ByVal asProc As String, ByVal asSpcno As String, Optional ByVal PID As String, Optional ByVal PWD As String, _
                              Optional ByVal FDate As String, Optional ByVal TDate As String, Optional ByVal TestS As String) As String
    
    Dim sProc As String
    Dim sParam As String
    
    Select_Param = ""
    sParam = ""
    sProc = asProc
    
    Select Case sProc
    Case gXml_S01, gXml_S02, gXml_S03, gXml_S07, gXml_S10
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>"
    
'Public Const gXml_BAR_SELECT = "PKG_LAB.INTERFACE_S29"      '바코드
'Public Const gXml_DAY_SELECT = "PKG_LAB.INTERFACE_S28"      'ONE Day
'Public Const gXml_TERM_SELECT = "PKG_LAB.INTERFACE_S23"     'From~To Date

    
    Case gXml_DAY_SELECT
        sParam = ""
        sParam = sParam & "<?xml version='1.0' encoding='UTF-8'?>"
        sParam = sParam & "<Table>"
    
    Case gXml_TERM_SELECT
        sParam = ""
        sParam = sParam & "<?xml version='1.0' encoding='UTF-8'?>" & vbCrLf
        sParam = sParam & "<NewDataSet>" & vbCrLf
        sParam = sParam & "<Table>" & vbCrLf
        sParam = sParam & "<QID><![CDATA[" & sProc & "]]></QID>" & vbCrLf
        sParam = sParam & "<QTYPE><![CDATA[Package]]></QTYPE>" & vbCrLf
        sParam = sParam & "<USERID><![CDATA[" & gHOSP.USERID & "]]></USERID>" & vbCrLf
        sParam = sParam & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & vbCrLf
        sParam = sParam & "<TABLENAME><![CDATA[]]></TABLENAME>" & vbCrLf
        sParam = sParam & "<P0><![CDATA[" & FDate & "]]></P0>" & vbCrLf
        sParam = sParam & "<P1><![CDATA[" & TDate & "]]></P1>" & vbCrLf
        sParam = sParam & "<P2><![CDATA[" & gHOSP.HOSPCD & "]]></P2>" & vbCrLf
        sParam = sParam & "<P3><![CDATA[" & gHOSP.PARTCD & "]]></P3>" & vbCrLf
        sParam = sParam & "<P4><![CDATA[" & TestS & "]]></P4>" & vbCrLf
        sParam = sParam & "<P5><![CDATA[0]]></P5>" & vbCrLf
        sParam = sParam & "<P6><![CDATA[9999]]></P6>" & vbCrLf
        If frmMain.ChkSave1.Value = "1" Then
            sParam = sParam & "<P7><![CDATA[1,2,3]]></P7>" & vbCrLf
        Else
            sParam = sParam & "<P7><![CDATA[1,2]]></P7>" & vbCrLf
        End If
        sParam = sParam & "<P8><![CDATA[]]></P8>" & vbCrLf
        sParam = sParam & "</Table>" & vbCrLf
        sParam = sParam & "</NewDataSet>" & vbCrLf
    
      
    Case gXml_ORDER_SELECT, gXml_BAR_SELECT
        sParam = ""
        sParam = sParam & "<?xml version='1.0' encoding='UTF-8'?>" & vbCrLf
        sParam = sParam & "<NewDataSet>" & vbCrLf
        sParam = sParam & "<Table>" & vbCrLf
        sParam = sParam & "<QID><![CDATA[" & sProc & "]]></QID>" & vbCrLf
        sParam = sParam & "<QTYPE><![CDATA[Package]]></QTYPE>" & vbCrLf
        sParam = sParam & "<USERID><![CDATA[" & gHOSP.USERID & "]]></USERID>" & vbCrLf
        sParam = sParam & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & vbCrLf
        sParam = sParam & "<TABLENAME><![CDATA[]]></TABLENAME>" & vbCrLf
        sParam = sParam & "<P0><![CDATA[" & asSpcno & "]]></P0>" & vbCrLf
        sParam = sParam & "<P1><![CDATA[1,2]]></P1>" & vbCrLf
        sParam = sParam & "<P2><![CDATA[]]></P2>" & vbCrLf
        sParam = sParam & "</Table>" & vbCrLf
        sParam = sParam & "</NewDataSet>" & vbCrLf
        
        
'<Table><QID><![CDATA[PKG_LAB.INTERFACE_S32]]></QID>
'<QTYPE><![CDATA[Package]]></QTYPE>
'<USERID><![CDATA[HIS]]></USERID>
'<EXECTYPE><![CDATA[FILL]]></EXECTYPE>
'<TABLENAME><![CDATA[]]></TABLENAME>
'<P0><![CDATA[20053002519]]></P0>
'<P1><![CDATA[]]></P1>
'</Table>
'
'<Table><QID><![CDATA[PKG_LAB.INTERFACE_S02]]></QID>
'<QTYPE><![CDATA[Package]]></QTYPE>
'<USERID><![CDATA[HIS]]></USERID>
'<EXECTYPE><![CDATA[FILL]]></EXECTYPE>
'<TABLENAME><![CDATA[]]></TABLENAME>
'<P0><![CDATA[20053100939]]></P0>
'<P1><![CDATA[]]></P1>
'</Table>

        
    Case gXml_LOGIN
        sParam = ""
        sParam = sParam & "<?xml version='1.0' encoding='UTF-8'?>"
        sParam = sParam & "<Table>"
        sParam = sParam & "<QID><![CDATA[" & sProc & "]]></QID>"
        sParam = sParam & "<QTYPE><![CDATA[Package]]></QTYPE>"
        sParam = sParam & "<USERID><![CDATA[RTE]]></USERID>"
        sParam = sParam & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
        'sParam = sParam & "<P0><![CDATA[" & gHOSP.SITE & "]]></P0>"
        sParam = sParam & "<P0><![CDATA[" & gHOSP.HOSPCD & "]]></P0>"
        sParam = sParam & "<P1><![CDATA[" & PID & "]]></P1>"
        sParam = sParam & "<P2><![CDATA[" & PWD & "]]></P2>"
        sParam = sParam & "</Table>"
        
    End Select
    
    Select_Param = sParam
    
End Function

Private Function TLA_Param(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String
    Dim sProc As String
    Dim sParam As String
    
    TLA_Param = ""
    sProc = asProc
    
    Select Case sProc
    Case gXml_S04, gXml_S24
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asDate1 & "]]></P0>" & _
                 "<P1><![CDATA[" & asDate2 & "]]></P1>" & _
                 "<P2><![CDATA[]]></P2>" & _
                 "</Table>"
    Case gXml_U07
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asBarcode & "]]></P0>" & _
                 "<P1><![CDATA[" & asDate1 & "]]></P1>" & _
                 "<P2><![CDATA[" & asDate2 & "]]></P2>" & _
                 "</Table>"
    End Select
    
    TLA_Param = sParam
    
End Function


'XML File Parsing===========================================================================================================
Public Sub display_online_parsing_ExamCode(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
'            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "TST_CD"
                giIndex = giIndex + 1
                ReDim Preserve gExam_Select(giIndex)
                gExam_Select(giIndex).TST_CD = xNode.nodeValue
                gExam_Select(giIndex).TST_CNT = giIndex + 1
                If gOrderExam = "" Then
                    gOrderExam = "'" & gExam_Select(giIndex).TST_CD & "'"
                Else
                    gOrderExam = gOrderExam & ", '" & gExam_Select(giIndex).TST_CD & "'"
                End If
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_ExamCode xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub display_online_parsing_PatInfo(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2
    giIndex = 0
    
    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
''            Select Case xNode.parentNode.nodeName
''            Case "TST_CD"
''                gPatInfo_Select.TST_CD = xNode.nodeValue
''                gPatInfo_Select.OK = 1
''            Case "ACPT_DTETM"
''                gPatInfo_Select.ACPT_DTETM = xNode.nodeValue
''            Case "ACPTNO_1"
''                gPatInfo_Select.ACPTNO_1 = xNode.nodeValue
''            Case "AGE"
''                gPatInfo_Select.Age = xNode.nodeValue
'''                Exit Sub
''            Case "HSP_CLS"
''                gPatInfo_Select.HSP_CLS = xNode.nodeValue
''            Case "ORD_SITE"
''                gPatInfo_Select.ORD_SITE = xNode.nodeValue
''            Case "PT_NM"
''                gPatInfo_Select.PT_NM = xNode.nodeValue
''            Case "PT_NO"
''                gPatInfo_Select.PT_NO = xNode.nodeValue
''            Case "RERUN"
''                gPatInfo_Select.RERUN = xNode.nodeValue
''            Case "SEX"
''                gPatInfo_Select.Sex = xNode.nodeValue
''            Case "SPC_CD_1"
''                gPatInfo_Select.SPC_CD_1 = xNode.nodeValue
''            Case "TST_CLS"
''                gPatInfo_Select.TST_CLS = xNode.nodeValue
''            Case "TST_FRCT_CD"
''                gPatInfo_Select.TST_FRCT_CD = xNode.nodeValue
''            Case "TST_FRCT_CD1"
''                gPatInfo_Select.TST_FRCT_CD1 = xNode.nodeValue
''            Case "TST_NM"
''                gPatInfo_Select.TST_NM = xNode.nodeValue
''            End Select
              
'              {
'              <?xml version="1.0" encoding="ISO-8859-1"?>
'              <string xmlns="http://tempuri.org/"><NewDataSet>
'              <Table>
'              <SPCM_NO>검체번호</SPCM_NO>
'              <ACPT_DTM>접수일자</ACPT_DTM>
'              <EXM_ACPT_NO>접수번호</EXM_ACPT_NO>
'              <PT_NO>환자번호</PT_NO>
'              <PT_NM>환자이름</<PT_NO>
'              <SEX_TP_CD>성별</SEX_TP_CD>
'              <PT_BRDY_DT>생년월일</PT_BRDY_DT>
'              <PT_HME_DEPT_CD>환자진료과</PT_HME_DEPT_CD>
'              <WD_DEPT_CD>병동</WD_DEPT_CD>
'              <EXM_CD>검사코드</EXM_CD>
'              <TH1_SPCM_CD>검체코드</TH1_SPCM_CD>
'              <HR24_URN_EXM_TM>24시간소변검사시간</HR24_URN_EXM_TM>
'              <HR24_URN_EXM_VLM_CNTE>24시간소변검사부피내용</HR24_URN_EXM_VLM_CNTE>
'              <RPRN_EXM_CD>판넬코드</RPRN_EXM_CD>
'              <EXM_PRGR_STS_CD>결과상태값</EXM_PRGR_STS_CD>
'              </Table>
'              </NewDataSet></string>
'              }
            
            Select Case xNode.parentNode.nodeName
                Case "SPCM_NO":                 gPatInfo_Select.SPCM_NO = xNode.nodeValue                       '검체번호
                Case "ACPT_DTM":                gPatInfo_Select.ACPT_DTM = xNode.nodeValue                      '접수일자
                Case "EXM_ACPT_NO":             gPatInfo_Select.EXM_ACPT_NO = xNode.nodeValue                   '접수번호
                Case "PT_NO":                   gPatInfo_Select.PT_NO = xNode.nodeValue                         '환자번호
                Case "PT_NM":                   gPatInfo_Select.PT_NM = xNode.nodeValue                         '환자이름
                Case "SEX_TP_CD":               gPatInfo_Select.SEX_TP_CD = xNode.nodeValue                     '    성별
                Case "PT_BRDY_DT":              gPatInfo_Select.PT_BRDY_DT = xNode.nodeValue                    '생년월일
                Case "PT_HME_DEPT_CD":          gPatInfo_Select.PT_HME_DEPT_CD = xNode.nodeValue                '진료과
                Case "WD_DEPT_CD":              gPatInfo_Select.WD_DEPT_CD = xNode.nodeValue                    '병동
                Case "TH1_SPCM_CD":             gPatInfo_Select.TH1_SPCM_CD = xNode.nodeValue                   '검체코드
                Case "HR24_URN_EXM_TM":         gPatInfo_Select.HR24_URN_EXM_TM = xNode.nodeValue               '24Hour 소뱐검사시간
                Case "HR24_URN_EXM_VLM_CNTE":   gPatInfo_Select.HR24_URN_EXM_VLM_CNTE = xNode.nodeValue         '24Hour 소뱐검사부피내용
                Case "RPRN_EXM_CD":             gPatInfo_Select.RPRN_EXM_CD = xNode.nodeValue                   '판넬코드
                Case "EXM_PRGR_STS_CD":         gPatInfo_Select.EXM_PRGR_STS_CD = xNode.nodeValue               '결과상태값
                Case "EXM_CD":                  gPatInfo_Select.EXM_CD = xNode.nodeValue                        '검사코드
                    gPatInfo_Select.OK = 1
                    giIndex = giIndex + 1
                    ReDim Preserve gExam_Select(giIndex)
                    
                    gExam_Select(giIndex).TST_CD = xNode.nodeValue
                    gExam_Select(giIndex).TST_CNT = giIndex + 1
                    
                    If gPatOrdCd = "" Then
                        'gOrderExam = "'" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & "'" & gExam_Select(giIndex).TST_CD & "',"
                    Else
                        'gOrderExam = gOrderExam & ", '" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & ", '" & gExam_Select(giIndex).TST_CD & "'"
                    End If

                    gPatTest(giIndex) = gExam_Select(giIndex).TST_CD

            End Select
        
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_PatInfo xNode.childNodes, Indent
        End If
    Next xNode
    
    
End Sub

Public Sub display_online_PatInfo(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    Dim i               As Integer
    Dim strTemp         As String
    Dim strAtbName      As String
    Dim strAtbValue     As String
    
    Dim xNode As MSXML2.IXMLDOMNode
'    Indent = Indent + 2
    giIndex = 0
    
'    For Each xNode In Nodes
    
'        If xNode.nodeType = 4 Then
''            Select Case xNode.parentNode.nodeName
''            Case "TST_CD"
''                gPatInfo_Select.TST_CD = xNode.nodeValue
''                gPatInfo_Select.OK = 1
''            Case "ACPT_DTETM"
''                gPatInfo_Select.ACPT_DTETM = xNode.nodeValue
''            Case "ACPTNO_1"
''                gPatInfo_Select.ACPTNO_1 = xNode.nodeValue
''            Case "AGE"
''                gPatInfo_Select.Age = xNode.nodeValue
'''                Exit Sub
''            Case "HSP_CLS"
''                gPatInfo_Select.HSP_CLS = xNode.nodeValue
''            Case "ORD_SITE"
''                gPatInfo_Select.ORD_SITE = xNode.nodeValue
''            Case "PT_NM"
''                gPatInfo_Select.PT_NM = xNode.nodeValue
''            Case "PT_NO"
''                gPatInfo_Select.PT_NO = xNode.nodeValue
''            Case "RERUN"
''                gPatInfo_Select.RERUN = xNode.nodeValue
''            Case "SEX"
''                gPatInfo_Select.Sex = xNode.nodeValue
''            Case "SPC_CD_1"
''                gPatInfo_Select.SPC_CD_1 = xNode.nodeValue
''            Case "TST_CLS"
''                gPatInfo_Select.TST_CLS = xNode.nodeValue
''            Case "TST_FRCT_CD"
''                gPatInfo_Select.TST_FRCT_CD = xNode.nodeValue
''            Case "TST_FRCT_CD1"
''                gPatInfo_Select.TST_FRCT_CD1 = xNode.nodeValue
''            Case "TST_NM"
''                gPatInfo_Select.TST_NM = xNode.nodeValue
''            End Select
              
'              {
'              <?xml version="1.0" encoding="ISO-8859-1"?>
'              <string xmlns="http://tempuri.org/"><NewDataSet>
'              <Table>
'              <SPCM_NO>검체번호</SPCM_NO>
'              <ACPT_DTM>접수일자</ACPT_DTM>
'              <EXM_ACPT_NO>접수번호</EXM_ACPT_NO>
'              <PT_NO>환자번호</PT_NO>
'              <PT_NM>환자이름</<PT_NO>
'              <SEX_TP_CD>성별</SEX_TP_CD>
'              <PT_BRDY_DT>생년월일</PT_BRDY_DT>
'              <PT_HME_DEPT_CD>환자진료과</PT_HME_DEPT_CD>
'              <WD_DEPT_CD>병동</WD_DEPT_CD>
'              <EXM_CD>검사코드</EXM_CD>
'              <TH1_SPCM_CD>검체코드</TH1_SPCM_CD>
'              <HR24_URN_EXM_TM>24시간소변검사시간</HR24_URN_EXM_TM>
'              <HR24_URN_EXM_VLM_CNTE>24시간소변검사부피내용</HR24_URN_EXM_VLM_CNTE>
'              <RPRN_EXM_CD>판넬코드</RPRN_EXM_CD>
'              <EXM_PRGR_STS_CD>결과상태값</EXM_PRGR_STS_CD>
'              </Table>
'              </NewDataSet></string>
'              }
                    
        For i = 3 To UBound(strRecvData)
                        
            strTemp = mGetP(strRecvData(i), 2, "<")
            strAtbName = mGetP(strTemp, 1, ">")
            strAtbValue = mGetP(strTemp, 2, ">")
            
            Select Case strAtbName
                Case "SPCM_NO":                 gPatInfo_Select.SPCM_NO = strAtbValue                       '검체번호
                Case "ACPT_DTM":                gPatInfo_Select.ACPT_DTM = strAtbValue                      '접수일자
                Case "EXM_ACPT_NO":             gPatInfo_Select.EXM_ACPT_NO = strAtbValue                   '접수번호
                Case "PT_NO":                   gPatInfo_Select.PT_NO = strAtbValue                         '환자번호
                Case "PT_NM":                   gPatInfo_Select.PT_NM = strAtbValue                         '환자이름
                Case "SEX_TP_CD":               gPatInfo_Select.SEX_TP_CD = strAtbValue                     '    성별
                Case "PT_BRDY_DT":              gPatInfo_Select.PT_BRDY_DT = strAtbValue                    '생년월일
                Case "PT_HME_DEPT_CD":          gPatInfo_Select.PT_HME_DEPT_CD = strAtbValue                '진료과
                Case "WD_DEPT_CD":              gPatInfo_Select.WD_DEPT_CD = strAtbValue                    '병동
                Case "TH1_SPCM_CD":             gPatInfo_Select.TH1_SPCM_CD = strAtbValue                   '검체코드
                Case "HR24_URN_EXM_TM":         gPatInfo_Select.HR24_URN_EXM_TM = strAtbValue               '24Hour 소뱐검사시간
                Case "HR24_URN_EXM_VLM_CNTE":   gPatInfo_Select.HR24_URN_EXM_VLM_CNTE = strAtbValue         '24Hour 소뱐검사부피내용
                Case "RPRN_EXM_CD":             gPatInfo_Select.RPRN_EXM_CD = strAtbValue                   '판넬코드
                Case "EXM_PRGR_STS_CD":         gPatInfo_Select.EXM_PRGR_STS_CD = strAtbValue               '결과상태값
                Case "EXM_CD":                  gPatInfo_Select.EXM_CD = strAtbValue                        '검사코드
                    gPatInfo_Select.OK = 1
                    giIndex = giIndex + 1
                    ReDim Preserve gExam_Select(giIndex)
                    ReDim Preserve gPatTest(giIndex)
                    
                    gExam_Select(giIndex).TST_CD = strAtbValue
                    gExam_Select(giIndex).TST_CNT = giIndex + 1
                    
                    If gPatOrdCd = "" Then
                        'gOrderExam = "'" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & "'" & gExam_Select(giIndex).TST_CD & "',"
                    Else
                        'gOrderExam = gOrderExam & ", '" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & ", '" & gExam_Select(giIndex).TST_CD & "'"
                    End If

                    gPatTest(giIndex) = gExam_Select(giIndex).TST_CD

            End Select
        Next
        'End If
        'If xNode.hasChildNodes Then
        '    display_online_parsing_PatInfo xNode.childNodes, Indent
        'End If
    'Next xNode
    
    

End Sub

Public Sub display_online_PatInfo_Term(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    Dim i               As Integer
    Dim strTemp         As String
    Dim strAtbName      As String
    Dim strAtbValue     As String
    
    Dim xNode As MSXML2.IXMLDOMNode
    giIndex = 0
    
    For i = 3 To UBound(strRecvData)
                    
        strTemp = mGetP(strRecvData(i), 2, "<")
        strAtbName = mGetP(strTemp, 1, ">")
        strAtbValue = mGetP(strTemp, 2, ">")
        
'    ACPT_DTE            As String  ' 검사일자
'    ACPTNO_1            As String  ' 접수번호
'    PT_NO               As String  ' 환자번호
'    PATNAME             As String  ' 환자명
'    SEX                 As String  ' 성별
'    AGE                 As String  ' 나이
'    SPC_NO              As String  ' 검체번호
'    ORD_SITE            As String  ' 발행처
'    WRK_UNT_CD          As String  ' 작업코드
'    TST_CD              As String  ' 검사코드
'    TST_SNM             As String  ' 검사명
'    SPC_NM              As String  ' 검체명
'    TST_STAT            As String  ' 상태\
'    ACPT_DTETM          As String  ' 접수일시
'    DOCTORNOTE_YN       As String  ' DOCTORNOTE_YN
'    SND_ARVL_NO_CNTE    As String  ' SND_ARVL_NO_CNTE
'    DEXM_YENU           As String  ' 연번호
'    OK                  As Integer

        Select Case strAtbName
            Case "ACPT_DTE":            gPInfo_T_Sel.ACPT_DTE = strAtbValue
            Case "ACPTNO_1":            gPInfo_T_Sel.ACPTNO_1 = strAtbValue
            Case "PT_NO":               gPInfo_T_Sel.PT_NO = strAtbValue
            Case "PATNAME":             gPInfo_T_Sel.PATNAME = strAtbValue
            Case "SEX":                 gPInfo_T_Sel.SEX = strAtbValue
            Case "AGE":                 gPInfo_T_Sel.AGE = strAtbValue
            Case "SPC_NO":              gPInfo_T_Sel.SPC_NO = strAtbValue
            Case "ORD_SITE":            gPInfo_T_Sel.ORD_SITE = strAtbValue
            Case "WRK_UNT_CD":          gPInfo_T_Sel.WRK_UNT_CD = strAtbValue
            Case "TST_CD":              gPInfo_T_Sel.TST_CD = strAtbValue
            Case "TST_SNM":             gPInfo_T_Sel.TST_SNM = strAtbValue
            Case "SPC_NM":              gPInfo_T_Sel.SPC_NM = strAtbValue
            Case "TST_STAT":            gPInfo_T_Sel.TST_STAT = strAtbValue
            Case "ACPT_DTETM":          gPInfo_T_Sel.ACPT_DTE = strAtbValue
            Case "DOCTORNOTE_YN":       gPInfo_T_Sel.DOCTORNOTE_YN = strAtbValue
            Case "SND_ARVL_NO_CNTE":    gPInfo_T_Sel.SND_ARVL_NO_CNTE = strAtbValue
            Case "DEXM_YENU":           gPInfo_T_Sel.DEXM_YENU = strAtbValue
            
                gPInfo_T_Sel.OK = 1
                giIndex = giIndex + 1
                ReDim Preserve gExam_Select(giIndex)
                ReDim Preserve gPatTest(giIndex)
                
                gExam_Select(giIndex).TST_CD = strAtbValue
                gExam_Select(giIndex).TST_CNT = giIndex + 1
                
                If gPatOrdCd = "" Then
                    gPatOrdCd = gPatOrdCd & "'" & gExam_Select(giIndex).TST_CD & "',"
                Else
                    gPatOrdCd = gPatOrdCd & ", '" & gExam_Select(giIndex).TST_CD & "'"
                End If

                gPatTest(giIndex) = gExam_Select(giIndex).TST_CD

        End Select
    Next

End Sub


Public Sub display_online_parsing_Login(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2
    giIndex = 0
    
    For Each xNode In Nodes
    
        If xNode.nodeType = 1 Then

            Select Case xNode.parentNode.nodeName
                Case "SPCM_NO":                 gPatInfo_Select.SPCM_NO = xNode.nodeValue                       '검체번호
                Case "ACPT_DTM":                gPatInfo_Select.ACPT_DTM = xNode.nodeValue                      '접수일자
                Case "EXM_ACPT_NO":             gPatInfo_Select.EXM_ACPT_NO = xNode.nodeValue                   '접수번호
                Case "PT_NO":                   gPatInfo_Select.PT_NO = xNode.nodeValue                         '환자번호
                Case "PT_NM":                   gPatInfo_Select.PT_NM = xNode.nodeValue                         '환자이름
                Case "SEX_TP_CD":               gPatInfo_Select.SEX_TP_CD = xNode.nodeValue                     '    성별
                Case "PT_BRDY_DT":              gPatInfo_Select.PT_BRDY_DT = xNode.nodeValue                    '생년월일
                Case "PT_HME_DEPT_CD":          gPatInfo_Select.PT_HME_DEPT_CD = xNode.nodeValue                '진료과
                Case "WD_DEPT_CD":              gPatInfo_Select.WD_DEPT_CD = xNode.nodeValue                    '병동
                Case "TH1_SPCM_CD":             gPatInfo_Select.TH1_SPCM_CD = xNode.nodeValue                   '검체코드
                Case "HR24_URN_EXM_TM":         gPatInfo_Select.HR24_URN_EXM_TM = xNode.nodeValue               '24Hour 소뱐검사시간
                Case "HR24_URN_EXM_VLM_CNTE":   gPatInfo_Select.HR24_URN_EXM_VLM_CNTE = xNode.nodeValue         '24Hour 소뱐검사부피내용
                Case "RPRN_EXM_CD":             gPatInfo_Select.RPRN_EXM_CD = xNode.nodeValue                   '판넬코드
                Case "EXM_PRGR_STS_CD":         gPatInfo_Select.EXM_PRGR_STS_CD = xNode.nodeValue               '결과상태값
                Case "EXM_CD":                  gPatInfo_Select.EXM_CD = xNode.nodeValue                        '검사코드
                    gPatInfo_Select.OK = 1
                    giIndex = giIndex + 1
                    ReDim Preserve gExam_Select(giIndex)
                    
                    gExam_Select(giIndex).TST_CD = xNode.nodeValue
                    gExam_Select(giIndex).TST_CNT = giIndex + 1
                    
                    If gPatOrdCd = "" Then
                        'gOrderExam = "'" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & "'" & gExam_Select(giIndex).TST_CD & "',"
                    Else
                        'gOrderExam = gOrderExam & ", '" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & ", '" & gExam_Select(giIndex).TST_CD & "'"
                    End If

                    gPatTest(giIndex) = gExam_Select(giIndex).TST_CD

            End Select
        
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_Login xNode.childNodes, Indent
        End If
    Next xNode
End Sub


Public Sub display_online_parsing_User(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "NM"
                gIFName = xNode.nodeValue
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_User xNode.childNodes, Indent
        End If
    Next xNode
End Sub


Public Sub display_online_parsing_TLAInfo(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
'            Select Case xNode.parentNode.nodeName
'            Case "TST_CD"
'                gPatInfo_Select.TST_CD = xNode.nodeValue
'                gPatInfo_Select.OK = 1
'            Case "ACPT_DTETM"
'                gPatInfo_Select.ACPT_DTETM = xNode.nodeValue
'            Case "ACPTNO_1"
'                gPatInfo_Select.ACPTNO_1 = xNode.nodeValue
'            Case "AGE"
'                gPatInfo_Select.AGE = xNode.nodeValue
''                Exit Sub
'            Case "HSP_CLS"
'                gPatInfo_Select.HSP_CLS = xNode.nodeValue
'            Case "ORD_SITE"
'                gPatInfo_Select.ORD_SITE = xNode.nodeValue
'            Case "PT_NM"
'                gPatInfo_Select.PT_NM = xNode.nodeValue
'            Case "PT_NO"
'                gPatInfo_Select.PT_NO = xNode.nodeValue
'            Case "RERUN"
'                gPatInfo_Select.RERUN = xNode.nodeValue
'            Case "SEX"
'                gPatInfo_Select.SEX = xNode.nodeValue
'            Case "SPC_CD_1"
'                gPatInfo_Select.SPC_CD_1 = xNode.nodeValue
'            Case "TST_CLS"
'                gPatInfo_Select.TST_CLS = xNode.nodeValue
'            Case "TST_FRCT_CD"
'                gPatInfo_Select.TST_FRCT_CD = xNode.nodeValue
'            Case "TST_FRCT_CD1"
'                gPatInfo_Select.TST_FRCT_CD1 = xNode.nodeValue
'            Case "TST_NM"
'                gPatInfo_Select.TST_NM = xNode.nodeValue
'            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_TLAInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub Display_Online_Parsing(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab

        End If
        If xNode.hasChildNodes Then
            display_online_parsing_TLAInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub



'===========================================================================================================XML File Parsing

'Result Trans sub start ====================================================================================================
Public Function Online_Result(ByVal asParam As String) As String

    Dim sRetStr As String


    Online_Result = ""

    gOnline_Ret = ""

    sRetStr = Online_Result_Qry(asParam)

    'SaveXMLFile sRetStr
    Xml_Log sRetStr, "res"

    Dim xDoc As MSXML2.DOMDocument

    Set xDoc = New MSXML2.DOMDocument

    If xDoc.Load(App.PATH & "\Res\res.xml") Then
    'If xDoc.Load(sRetStr) Then
        ' 문서가 성공적으로 로드되었습니다.
        ' 이제 재미있는 작업을 수행합니다.
        Display_Online_Parsing xDoc.childNodes, 0
    Else
        ' 문서를 로드하지 못했습니다.
        Dim strErrText As String
        Dim xPE As MSXML2.IXMLDOMParseError
       ' ParseError 개체를 가져옵니다
        Set xPE = xDoc.parseError
        With xPE
            strErrText = "Your XML Document failed to load" & _
                         "due the following error." & vbCrLf & _
                         "Error #: " & .errorCode & ": " & xPE.reason & _
                         "Line #: " & .Line & vbCrLf & _
                         "Line Position: " & .linepos & vbCrLf & _
                         "Position In File: " & .filepos & vbCrLf & _
                         "Source Text: " & .srcText & vbCrLf & _
                         "Document URL: " & .URL
        End With

        'SaveData strErrText
    End If

    Set xPE = Nothing

    Set xDoc = Nothing

    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_Result = LEFT(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If

End Function

Public Function Online_Result_Qry(ByVal asParam As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String

    On Error GoTo ErrHandle

    Set oSOAP = New MSSOAPLib30.SoapClient30

    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit gHOSP.APIURL & "?wsdl"
    
    strDiv = gXml_RESULT_UPLOAD
    
    sParam = asParam
    
    SaveXML_Data "[Save Result]" & sParam
    
    'send = oSOAP.LMService("SETQUERY", sParam)
    
    '예제) this.WcfService.ServiceReturnCustomType(sType, "", "XML", "N", strXML);
    send = oSOAP.ServiceReturnCustomType("SETQUERY", "", "XML", "N", sParam)
    
    SaveXML_Data "[Save Result => Return]" & send
    Online_Result_Qry = send
    Set oSOAP = Nothing
    DoEvents
    Exit Function

ErrHandle:
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
End Function

'==================================================================================================== Result Trans sub start

'데이터 저장=================================================================================================================
Public Sub SaveXML_Data(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum

    FilNum = FreeFile

    If Dir(App.PATH & "\" & "XML", vbDirectory) <> "XML" Then
        MkDir (App.PATH & "\XML")
    End If

    Open App.PATH & "\XML" & "\" & Date & ".log" For Append As FilNum
    Print #FilNum, Time & " " & argSQL
    Close FilNum
End Sub

Public Sub Xml_Log(argSQL As String, argFileName As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.PATH & "\" & "XML", vbDirectory) <> "XML" Then
        MkDir (App.PATH & "\" & "XML")
    End If
    
    sFileName = argFileName
    If Dir(App.PATH & "\" & "XML" & "\" & sFileName & ".xml") <> "" Then
        Kill App.PATH & "\" & "XML" & "\" & sFileName & ".xml"
    End If
    
    Open App.PATH & "\" & "XML" & "\" & sFileName & ".xml" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

'=================================================================================================================데이터 저장



