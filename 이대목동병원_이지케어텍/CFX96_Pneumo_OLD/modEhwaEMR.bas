Attribute VB_Name = "modEhwaEMR"
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
Public Const gXml_ORDER_SELECT = "PKG_MSE_LM_INTERFACE.PC_MSE_ORDER_SELECT"
'-- 결과저장
Public Const gXml_RESULT_UPLOAD = "PKG_MSE_LM_INTERFACE.PC_MSE_INTERFACE_SAVE"

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
    Sex         As String
    Age         As String
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
        Online_TLA = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If
    
End Function

Public Function Online_XML(ByVal asProc As String, ByVal asSpcno As String, Optional ByVal asDIV As String, Optional ByVal PID As String, Optional ByVal PWD As String) As String
    Dim xmlDoc        As MSXML2.DOMDocument30
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
'    Dim MsgType As String
'    Dim strBuffer As String
'    Dim intRow As Long
'    Dim varBuffer As Variant
'    Dim blnQc     As Boolean
'    Dim i, J, k, m As Integer
'    Dim ii, jj, kk  As Integer
'    Dim strOData    As String
'    Dim strRData    As String
'
    Dim i As Integer
    Dim j As Integer
    
    
    Online_XML = ""
    sFileName = "Res"
    
    '파라미터를 만든다
    sParam = Select_Param(asProc, asSpcno, PID, PWD)
    
    '파라미터를 서버에 전송항여 리턴값을 받아온다.
    sRetStr = Online_XML_Qry(asDIV, sParam)
    
    'sRetStr = "<?xml version='1.0' encoding='UTF-8'?>" & vbCrLf & sRetStr
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    
    Dim strData As String
    
    Erase strRecvData
    intBufCnt = 1000
    ReDim Preserve strRecvData(1)
    j = 1
    For i = 1 To Len(sRetStr)
        strData = Mid(sRetStr, i, 1)
        Select Case strData
            'Case "<"
            '    strRecvData(j) = strRecvData(j) & strData
            'Case "/"
            Case vbCr
            Case vbLf
                'strRecvData(j) = strRecvData(j) & strData
                j = j + 1
                ReDim Preserve strRecvData(j)
            Case Else
                
                strRecvData(j) = strRecvData(j) & strData
        End Select
    Next
    Set xmlDoc = New MSXML2.DOMDocument30
    
    
    'If xmlDoc.loadXML(App.PATH & "\XML\" & sFileName & ".xml") Then
    'If xmlDoc.loadXML(sRetStr) Then
    If UBound(strRecvData) > 1 Then
        
    '    Set Nodes = xmlDoc.documentElement
        
    '    If xmlDoc.readyState = 4 And xmlDoc.parseError.errorCode = 0 Then
    '        Online_XML = Nodes.Text
    '    Else
        
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
                display_online_PatInfo xmlDoc.childNodes, 0
                Online_XML = strRecvData(3)
            
            End Select
    '        Display_Online_Parsing_Test xmlDoc.childNodes, 0
            
     '   End If
    
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
        Online_XML = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If
    
End Function

Public Function Online_XML_Qry(ByVal asStrDiv As String, ByVal asParam As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    oSOAP.MSSoapInit gHOSP.APIURL & "?wsdl"
    strDiv = asStrDiv
    sParam = asParam
    
    SaveXML_Data "[Use Proc => " & strDiv & " ]" & sParam
    send = oSOAP.LMService(strDiv, sParam)
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

Private Function Select_Param(ByVal asProc As String, ByVal asSpcno As String, Optional ByVal PID As String, Optional ByVal PWD As String) As String
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
    
    Case gXml_ORDER_SELECT
        sParam = ""
        sParam = sParam & "<?xml version='1.0' encoding='UTF-8'?>"
        sParam = sParam & "<Table>"
        sParam = sParam & "<QID><![CDATA[" & sProc & "]]></QID>"
        sParam = sParam & "<QTYPE><![CDATA[Package]]></QTYPE>"
        sParam = sParam & "<USERID><![CDATA[RTE]]></USERID>"
        sParam = sParam & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
        sParam = sParam & "<P0><![CDATA[" & mResult.SITE & "]]></P0>"
        sParam = sParam & "<P1><![CDATA[" & gHOSP.MACHCD & "]]></P1>"
        sParam = sParam & "<P2><![CDATA[]]></P2>"
        sParam = sParam & "<P3><![CDATA[]]></P3>"
        sParam = sParam & "<P4><![CDATA[]]></P4>"
        sParam = sParam & "<P5><![CDATA[]]></P5>"
        sParam = sParam & "<P6><![CDATA[]]></P6>"
        sParam = sParam & "<P7><![CDATA[" & asSpcno & "]]></P7>"
        sParam = sParam & "<P8><![CDATA[]]></P8>"
        sParam = sParam & "<P9><![CDATA[]]></P9>"
        sParam = sParam & "<P10><![CDATA[]]></P10>"
        sParam = sParam & "<P11><![CDATA[]]></P11>"
        sParam = sParam & "<P12><![CDATA[]]></P12>"
        sParam = sParam & "<P13><![CDATA[]]></P13>"
        sParam = sParam & "<P14><![CDATA[]]></P14>"
        sParam = sParam & "<P15><![CDATA[]]></P15>"
        sParam = sParam & "<P16><![CDATA[]]></P16>"
        sParam = sParam & "<P17><![CDATA[]]></P17>"
        sParam = sParam & "<P18><![CDATA[]]></P18>"
        sParam = sParam & "</Table>"
        
'<?xml version='1.0' encoding='UTF-8'?>
'<Table>
'<QID>
'<![CDATA[PKG_MSE_LM_INTERFACE.PC_MSE_USER_SELECT]]>
'</QID>
'<QTYPE>
'<![CDATA[Package]]>
'</QTYPE>
'<USERID>
'<![CDATA[RTE]]>
'</USERID>
'<EXECTYPE>
'<![CDATA[FILL]]>
'</EXECTYPE>
'<P0>
'<![CDATA[02]]>
'</P0>
'<P1>
'<![CDATA[C0EMR]]>
'</P1>
'<P2>
'<![CDATA[11111]]>
'</P2>
'</Table>
        
    Case gXml_LOGIN
        sParam = ""
        sParam = sParam & "<?xml version='1.0' encoding='UTF-8'?>"
        sParam = sParam & "<Table>"
        sParam = sParam & "<QID><![CDATA[" & sProc & "]]></QID>"
        sParam = sParam & "<QTYPE><![CDATA[Package]]></QTYPE>"
        sParam = sParam & "<USERID><![CDATA[RTE]]></USERID>"
        sParam = sParam & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
        sParam = sParam & "<P0><![CDATA[" & gHOSP.SITE & "]]></P0>"
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

        SaveData strErrText
    End If

    Set xPE = Nothing

    Set xDoc = Nothing

    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_Result = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
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
    
    strDiv = gXml_RESULT_UPLOAD 'PKG_MSE_LM_INTERFACE.PC_MSE_INTERFACE_SAVE

    
    sParam = asParam
    
    SaveXML_Data "[Save Result]" & sParam
    'send = oSOAP.wsLISInterface(strDiv, sParam)
    send = oSOAP.LMService("SETQUERY", sParam)
    
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



