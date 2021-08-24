Attribute VB_Name = "nnc_emr"
Option Explicit

Public gOnline_Test As String
Public gServerPath As String
Public giIndex  As Long

Public Const gXml_S01 = "PG_SRL.SLP91_S01"
Public Const gXml_S02 = "PG_SRL.SLP91_S02"
Public Const gXml_S03 = "PG_SRL.SLP91_S03"

Public Const gXml_S04 = "PG_SRL.SLP91_S04"
Public Const gXml_S05 = "PG_SRL.SLP91_S05"
Public Const gXml_S07 = "PG_SRL.SLP91_S07"

Public Const gXml_S10 = "PG_SRL.SLP91_S10"
Public Const gXml_S24 = "PG_SRL.SLP91_S24"


Public Const gXml_U07 = "PG_SRL.SLP91_U01"

Type Exam_Select
    TST_CD      As String
    TST_CNT     As Integer
End Type

Type PatInfo_Select
    TST_CD      As String
    TST_NM      As String
    TST_FRCT_CD As String
    ACPTNO_1    As String
    PT_NO       As String
    PT_NM       As String
    SEX         As String
    SPC_CD_1    As String
    MEDDEPT    As String
    ORD_SITE    As String
    TST_CLS     As String
    RERUN       As String
    TST_FRCT_CD1    As String
    HSP_CLS     As String
    ACPT_DTETM  As String
    ACPT_DTE  As String
    
    AGE         As String
    OK          As Integer
End Type

Type Cobas4800_Result
    Barcode         As String
    spCode          As String
    Pos             As String
    Type16_Res      As String
    Type18_Res      As String
    TypeOther_Res   As String
    Type16_CT       As String
    Type18_CT       As String
    TypeOther_CT    As String
    Result          As String
End Type

Type Cobas_Barcode
    spCode  As String
    Barcode As String
End Type

Type Cobas4800_Pos
    Barcode         As String
    Pos             As String
End Type

Public gCobasBarcode    As Cobas_Barcode

Public gCobasPos    As Cobas4800_Pos
Public gCobas4800Res As Cobas4800_Result

Public gExam_Select() As Exam_Select
Public gPat_Info_Select As PatInfo_Select

Type TLAInfo_Select
    TST_DTE     As String
    SPCNO       As String
    OK          As Integer
End Type

Public gTLA_Info_Select() As TLAInfo_Select

Public gCobasFlag As Boolean
Public gXMLState As String

Public Sub Clear_Barcode()
    gCobasBarcode.Barcode = ""
    gCobasBarcode.spCode = ""
    
End Sub
Public Sub Clear_Cobas()
    gCobas4800Res.Barcode = ""
    gCobas4800Res.Result = ""
    gCobas4800Res.Type16_CT = ""
    gCobas4800Res.Type16_Res = ""
    gCobas4800Res.Type18_CT = ""
    gCobas4800Res.Type18_Res = ""
    gCobas4800Res.TypeOther_CT = ""
    gCobas4800Res.TypeOther_Res = ""
    gCobas4800Res.spCode = ""
    
End Sub
Public Sub Clear_XML_Exam()
    giIndex = -1
    ReDim gExam_Select(0)
End Sub

Public Sub Clear_XML_TLAInfo()
    giIndex = -1
    ReDim gTLA_Info_Select(0)
End Sub

Public Sub Clear_XML_PInfo()
    gPat_Info_Select.ACPT_DTETM = ""
    gPat_Info_Select.ACPTNO_1 = ""
    gPat_Info_Select.AGE = ""
    gPat_Info_Select.HSP_CLS = ""
    gPat_Info_Select.OK = -1
    gPat_Info_Select.ORD_SITE = ""
    gPat_Info_Select.PT_NM = ""
    gPat_Info_Select.PT_NO = ""
    gPat_Info_Select.RERUN = ""
    gPat_Info_Select.SEX = ""
    gPat_Info_Select.SPC_CD_1 = ""
    gPat_Info_Select.TST_CD = ""
    gPat_Info_Select.TST_CLS = ""
    gPat_Info_Select.TST_FRCT_CD = ""
    gPat_Info_Select.TST_FRCT_CD1 = ""
    gPat_Info_Select.TST_NM = ""
    
End Sub

Public Function Online_TLA(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    Online_TLA = ""
    sFileName = "Res"
    
    sParam = TLA_Param(asProc, asDate1, asDate2, asBarcode)
    
    sRetStr = Online_XML_Qry(asProc, sParam)
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    Dim xDoc As MSXML.DOMDocument
    Set xDoc = New MSXML.DOMDocument
    If xDoc.Load(App.Path & "\XML\" & sFileName & ".xml") Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_S04
            Clear_XML_TLAInfo
            display_online_parsing_TLAInfo xDoc.childNodes, 0
        Case gXml_S24
            gIFName = ""
            display_online_parsing_User xDoc.childNodes, 0
        End Select
'        Display_Online_Parsing_Test xDoc.childNodes, 0
    Else
        ' 문서를 로드하지 못했습니다.
        Dim strErrText As String
        Dim xPE As MSXML.IXMLDOMParseError
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
                         "Document URL: " & .url
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

Public Function Online_XML(ByVal asProc As String, ByVal asSpcno As String) As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    Online_XML = ""
    sFileName = "Res"
    
    sParam = Select_Param(asProc, asSpcno)
    
    sRetStr = Online_XML_Qry(asProc, sParam)
    
'''    sRetStr = "<?xml version='1.0' encoding='euc-kr'?>"
'''    sRetStr = sRetStr & chrLF & "<NewDataSet>"
'''    sRetStr = sRetStr & chrLF & "    <Table0>"
'''    sRetStr = sRetStr & chrLF & "        <TST_CD><![CDATA[L2725]]></TST_CD>"
'''    sRetStr = sRetStr & chrLF & "        <TST_NM><![CDATA[HPV(Real-time PCR)]]></TST_NM>"
'''    sRetStr = sRetStr & chrLF & "        <TST_FRCT_CD><![CDATA[L25]]></TST_FRCT_CD>"
'''    sRetStr = sRetStr & chrLF & "        <ACPTNO_1><![CDATA[1001]]></ACPTNO_1>"
'''    sRetStr = sRetStr & chrLF & "        <PT_NO><![CDATA[33216266]]></PT_NO>"
'''    sRetStr = sRetStr & chrLF & "        <PT_NM><![CDATA[김영순]]></PT_NM>"
'''    sRetStr = sRetStr & chrLF & "        <SEX><![CDATA[F]]></SEX>"
'''    sRetStr = sRetStr & chrLF & "        <SPC_CD_1><![CDATA[1SWA]]></SPC_CD_1>"
'''    sRetStr = sRetStr & chrLF & "        <ORD_SITE><![CDATA[PRP]]></ORD_SITE>"
'''    sRetStr = sRetStr & chrLF & "        <TST_CLS><![CDATA[N]]></TST_CLS>"
'''    sRetStr = sRetStr & chrLF & "        <RERUN><![CDATA[N]]></RERUN>"
'''    sRetStr = sRetStr & chrLF & "        <TST_FRCT_CD1><![CDATA[L25]]></TST_FRCT_CD1>"
'''    sRetStr = sRetStr & chrLF & "        <HSP_CLS><![CDATA[8]]></HSP_CLS>"
'''    sRetStr = sRetStr & chrLF & "        <ACPT_DTETM><![CDATA[2014-01-09 11:30:27]]></ACPT_DTETM>"
'''    sRetStr = sRetStr & chrLF & "        <AGE><![CDATA[76]]></AGE>"
'''    sRetStr = sRetStr & chrLF & "        <ACPT_DTE><![CDATA[20140109]]></ACPT_DTE>"
'''    sRetStr = sRetStr & chrLF & "        <PATSECT><![CDATA[O]]></PATSECT>"
'''    sRetStr = sRetStr & chrLF & "        <CAUTION_YN><![CDATA[N]]></CAUTION_YN>"
'''    sRetStr = sRetStr & chrLF & "        <MEDDEPT><![CDATA[PRP]]></MEDDEPT>"
'''    sRetStr = sRetStr & chrLF & "    </Table0>"
'''    sRetStr = sRetStr & chrLF & "</NewDataSet>"

    

    'SaveXMLFile sRetStr
'''    Xml_Log sRetStr, sFileName
    SaveXML_Data sRetStr
        
    Dim xDoc As MSXML.DOMDocument
    Set xDoc = New MSXML.DOMDocument
'''    If xDoc.Load(App.Path & "\XML\" & sFileName & ".xml") Then
    If xDoc.LoadXml(sRetStr) Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_S03
            Clear_XML_PInfo
            display_online_parsing_PatInfo xDoc.childNodes, 0
        Case gXml_S05, gXml_S07
            Clear_XML_Exam
            display_online_parsing_ExamCode xDoc.childNodes, 0
        End Select
'        Display_Online_Parsing_Test xDoc.childNodes, 0
    Else
        ' 문서를 로드하지 못했습니다.
        Dim strErrText As String
        Dim xPE As MSXML.IXMLDOMParseError
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
                         "Document URL: " & .url
        End With

        SaveXML_Data strErrText
    End If
    
    
    Set xPE = Nothing
    Set xDoc = Nothing
    
    Select Case asProc
    Case gXml_S03
        If gPat_Info_Select.OK = 1 Then
            Online_XML = 1
        Else
            Online_XML = 0
        End If
    Case gXml_S05, gXml_S07
        If gExam_Select(0).TST_CNT > 0 Then
            Online_XML = 1
        Else
            Online_XML = 0
        End If
    End Select
    
End Function

Public Function Online_XML_Qry(ByVal asStrDiv As String, ByVal asParam As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    oSOAP.MSSoapInit gServerPath
    strDiv = asStrDiv
    sParam = asParam

'''    SaveXML_Data "[Use Proc => " & strDiv & " ]" & sParam
    send = oSOAP.wsLISInterface(strDiv, sParam)
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

Private Function Select_Param(ByVal asProc As String, ByVal asSpcno As String) As String
    Dim sProc As String
    Dim sParam As String
    
    Select_Param = ""
    sProc = asProc
    
    Select Case sProc
    Case gXml_S01, gXml_S02, gXml_S03, gXml_S05, gXml_S07, gXml_S10
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>"
        
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

Public Sub display_online_parsing_User(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "NM"
                gIFName = xNode.nodeValue
            End Select
'            gTLA_Info_Select(giIndex).OK = 1
            
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_User xNode.childNodes, Indent
        End If
    Next xNode
End Sub



Public Sub display_online_parsing_ExamCode(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
'            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "TST_CD"
                giIndex = giIndex + 1
                ReDim Preserve gExam_Select(giIndex)
                gExam_Select(giIndex).TST_CD = xNode.nodeValue
                If gOrderExam = "" Then
                    gOrderExam = "'" & gExam_Select(giIndex).TST_CD & "'"
                Else
                    gOrderExam = gOrderExam & ", '" & gExam_Select(giIndex).TST_CD & "'"
                End If
                gExam_Select(giIndex).TST_CNT = giIndex + 1
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_ExamCode xNode.childNodes, Indent
        End If
    Next xNode
End Sub




Public Sub display_online_parsing_PatInfo(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2


    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            Select Case xNode.parentNode.nodeName
            Case "TST_CD"
                gPat_Info_Select.TST_CD = xNode.nodeValue
                gPat_Info_Select.OK = 1
                
                If gOrderExam = "" Then
                    gOrderExam = "'" & gPat_Info_Select.TST_CD & "'"
                Else
                    gOrderExam = gOrderExam & ", '" & gPat_Info_Select.TST_CD & "'"
                End If
                
            Case "ACPT_DTETM"
                gPat_Info_Select.ACPT_DTETM = xNode.nodeValue
            Case "ACPTNO_1"
                gPat_Info_Select.ACPTNO_1 = xNode.nodeValue
            Case "AGE"
                gPat_Info_Select.AGE = xNode.nodeValue
'                Exit Sub
            Case "HSP_CLS"
                gPat_Info_Select.HSP_CLS = xNode.nodeValue
            Case "ORD_SITE"
                gPat_Info_Select.ORD_SITE = xNode.nodeValue
            Case "PT_NM"
                gPat_Info_Select.PT_NM = xNode.nodeValue
            Case "PT_NO"
                gPat_Info_Select.PT_NO = xNode.nodeValue
            Case "RERUN"
                gPat_Info_Select.RERUN = xNode.nodeValue
            Case "SEX"
                gPat_Info_Select.SEX = xNode.nodeValue
            Case "SPC_CD_1"
                gPat_Info_Select.SPC_CD_1 = xNode.nodeValue
            Case "TST_CLS"
                gPat_Info_Select.TST_CLS = xNode.nodeValue
            Case "TST_FRCT_CD"
                gPat_Info_Select.TST_FRCT_CD = xNode.nodeValue
            Case "TST_FRCT_CD1"
                gPat_Info_Select.TST_FRCT_CD1 = xNode.nodeValue
            Case "TST_NM"
                gPat_Info_Select.TST_NM = xNode.nodeValue
            Case "MEDDEPT"
                gPat_Info_Select.MEDDEPT = xNode.nodeValue
            Case "ACPT_DTE"
                gPat_Info_Select.ACPT_DTE = xNode.nodeValue
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_PatInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub display_online_parsing_TLAInfo(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "TST_DTE"
                giIndex = giIndex + 1
                ReDim Preserve gTLA_Info_Select(giIndex)
                
                gTLA_Info_Select(giIndex).TST_DTE = xNode.nodeValue
            Case "SPCNO"
                gTLA_Info_Select(giIndex).SPCNO = xNode.nodeValue
            End Select
            gTLA_Info_Select(giIndex).OK = 1
            
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_TLAInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

'===========================================================================================================XML File Parsing

'Result Trans sub start ====================================================================================================
Public Function Online_Result(ByVal asParam As String) As String
'
'    Dim sRetStr As String
'
'
'    Online_Result = ""
'
'    gOnline_Ret = ""
'
'    sRetStr = Online_Result_Qry(asParam)
'
'    'SaveXMLFile sRetStr
'    Xml_Log sRetStr, "res"
'
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    If xDoc.Load(App.Path & "\Res\res.xml") Then
'    'If xDoc.Load(sRetStr) Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_Online_Parsing xDoc.childNodes, 0
'    Else
'        ' 문서를 로드하지 못했습니다.
'        Dim strErrText As String
'        Dim xPE As MSXML.IXMLDOMParseError
'       ' ParseError 개체를 가져옵니다
'        Set xPE = xDoc.parseError
'        With xPE
'            strErrText = "Your XML Document failed to load" & _
'                         "due the following error." & vbCrLf & _
'                         "Error #: " & .errorCode & ": " & xPE.reason & _
'                         "Line #: " & .Line & vbCrLf & _
'                         "Line Position: " & .linepos & vbCrLf & _
'                         "Position In File: " & .filepos & vbCrLf & _
'                         "Source Text: " & .srcText & vbCrLf & _
'                         "Document URL: " & .url
'        End With
'
'        SaveData strErrText
'    End If
'
'    Set xPE = Nothing
'
'    Set xDoc = Nothing
'
'    If InStr(1, gOnline_Ret, vbTab) > 0 Then
'        Online_Result = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
'    End If
'
End Function

Public Function Online_Result_Qry(ByVal asParam As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String

    On Error GoTo ErrHandle

    Set oSOAP = New MSSOAPLib30.SoapClient30

    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit gServerPath
    
    strDiv = "PG_SRL.SLP91_P03"
    
    sParam = asParam
    
    SaveXML_Data "[Save Result]" & sParam
    send = oSOAP.wsLISInterface(strDiv, sParam)
    SaveXML_Data "[Save Result => Return]" & send
    Online_Result_Qry = send
    Set oSOAP = Nothing
    DoEvents
    Exit Function

ErrHandle:
    If oSOAP.FaultString <> "" Then
        SaveXML_Data "[Save Error => Return]" & Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        SaveXML_Data "[Save Error => Return]" & Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
End Function

'==================================================================================================== Result Trans sub start

'데이터 저장=================================================================================================================
Public Sub SaveXML_Data(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum

    FilNum = FreeFile

    If Dir(App.Path & "\" & "XML", vbDirectory) <> "XML" Then
        MkDir (App.Path & "\XML")
    End If

    Open App.Path & "\XML" & "\" & Date & ".log" For Append As FilNum
    Print #FilNum, Time & " " & argSQL
    Close FilNum
End Sub

Public Sub Xml_Log(argSQL As String, argFileName As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\" & "XML", vbDirectory) <> "XML" Then
        MkDir (App.Path & "\" & "XML")
    End If
    
    sFileName = argFileName
    If Dir(App.Path & "\" & "XML" & "\" & sFileName & ".xml") <> "" Then
        Kill App.Path & "\" & "XML" & "\" & sFileName & ".xml"
    End If
    
    Open App.Path & "\" & "XML" & "\" & sFileName & ".xml" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

'=================================================================================================================데이터 저장
Public Function Cobas4800_Auto(ByVal asFileName As String) As Integer

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    
    Dim strMSG As String



    Cobas4800_Auto = -1
    
    Dim xDoc As MSXML.DOMDocument
'''    Dim node1 As IXMLDOMNode
'''    Dim node2 As IXMLDOMNodeList
    
    Set xDoc = New MSXML.DOMDocument
    If xDoc.LoadXml(frmInterface.txtXMLRes.Text) = True Then
'''        Set node2 = xDoc.selectNodes("//DomainObjects/TestOrders/TestOrder/TestResult/Details/StringValue")
''''''        strMSG = node2.Item
'''
'''        Set node1 = xDoc.selectSingleNode("//DomainObjects/TestOrders/TestOrder/TestResult/Details/StringValue")
'''        strMSG = node1.Attributes(0).nodeName & node1.Attributes(0).nodeValue & node1.Attributes(1).nodeName & node1.Attributes(1).nodeValue
'''        Set node1 = Nothing
'''        strMSG = ""
        Cobas4800_Xml_DisPlay xDoc.childNodes, 0

    Else
        ' 문서를 로드하지 못했습니다.
        Dim strErrText As String
        Dim xPE As MSXML.IXMLDOMParseError
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
                         "Document URL: " & .url
        End With

        SaveXML_Data strErrText
    End If
    
    
    Set xPE = Nothing
    Set xDoc = Nothing
    
    Cobas4800_Auto = 1
    
End Function

Public Function Cobas4800_Xml(ByVal asFileName As String) As Integer

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    
    Dim strMSG As String



    Cobas4800_Xml = -1
    
    Dim xDoc As MSXML.DOMDocument
'''    Dim node1 As IXMLDOMNode
'''    Dim node2 As IXMLDOMNodeList
    
    Set xDoc = New MSXML.DOMDocument
    If xDoc.Load(asFileName) = True Then
'''        Set node2 = xDoc.selectNodes("//DomainObjects/TestOrders/TestOrder/TestResult/Details/StringValue")
''''''        strMSG = node2.Item
'''
'''        Set node1 = xDoc.selectSingleNode("//DomainObjects/TestOrders/TestOrder/TestResult/Details/StringValue")
'''        strMSG = node1.Attributes(0).nodeName & node1.Attributes(0).nodeValue & node1.Attributes(1).nodeName & node1.Attributes(1).nodeValue
'''        Set node1 = Nothing
'''        strMSG = ""
        Cobas4800_Xml_DisPlay xDoc.childNodes, 0

    Else
        ' 문서를 로드하지 못했습니다.
        Dim strErrText As String
        Dim xPE As MSXML.IXMLDOMParseError
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
                         "Document URL: " & .url
        End With

        SaveXML_Data strErrText
    End If
    
    
    Set xPE = Nothing
    Set xDoc = Nothing
    
    Cobas4800_Xml = 1
    
End Function

Public Sub Cobas4800_Xml_DisPlay(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Dim xNodeList As MSXML.IXMLDOMNodeList
    Dim strRes As String
    Dim iRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim strTest As String
    Dim X As Integer
    
    Indent = Indent + 2
    For Each xNode In Nodes
        
'''        If xNode.nodeName = "Tests" Then Exit Sub
'''        If xNode.nodeName = "StringValue" Then
'''            strRes = ""
'''        End If

        
        If xNode.nodeType = 3 Then
            If xNode.parentNode.nodeName = "Interpretation" Then
                gCobas4800Res.Result = xNode.nodeValue
                
                iRow = frmInterface.vasXML.DataRowCnt + 1
                If frmInterface.vasXML.MaxRows < iRow Then
                    frmInterface.vasXML.MaxRows = iRow
                End If
                
                SetText frmInterface.vasXML, gCobas4800Res.Barcode, iRow, 1
                SetText frmInterface.vasXML, gCobas4800Res.Pos, iRow, 2
                SetText frmInterface.vasXML, gCobas4800Res.Type16_CT, iRow, 3
                SetText frmInterface.vasXML, gCobas4800Res.Type16_Res, iRow, 4
                SetText frmInterface.vasXML, gCobas4800Res.Type18_CT, iRow, 5
                SetText frmInterface.vasXML, gCobas4800Res.Type18_Res, iRow, 6
                SetText frmInterface.vasXML, gCobas4800Res.TypeOther_CT, iRow, 7
                SetText frmInterface.vasXML, gCobas4800Res.TypeOther_Res, iRow, 8
                SetText frmInterface.vasXML, gCobas4800Res.Result, iRow, 9
                SetText frmInterface.vasXML, gCobas4800Res.spCode, iRow, 10
            End If
            
            If xNode.parentNode.nodeName = "Position" Then
                gCobas4800Res.Pos = xNode.nodeValue
            End If
            
            If xNode.parentNode.nodeName = "Barcode" Then
                gCobasBarcode.Barcode = xNode.nodeValue
                
                For X = 1 To frmInterface.vasXML.DataRowCnt
                    If Trim(GetText(frmInterface.vasXML, X, 10)) = Trim(gCobasBarcode.spCode) Then
                        SetText frmInterface.vasXML, gCobasBarcode.Barcode, X, 1
                        Exit For
                    End If
                Next
            End If
            
        End If
        
        If xNode.nodeType = 1 Then
            Select Case xNode.nodeName
            Case "Sample"
                Clear_Barcode
                gCobasBarcode.spCode = xNode.Attributes(0).nodeValue
            Case "TestOrder"
                Clear_Cobas
                gCobas4800Res.spCode = xNode.Attributes(2).nodeValue
                
            Case "StringValue"
                Select Case xNode.Attributes(0).nodeValue
                
                Case "Ct:0" 'Other
                    gCobas4800Res.TypeOther_CT = xNode.Attributes(1).nodeValue
                Case "Ct:1" '16
                    gCobas4800Res.Type16_CT = xNode.Attributes(1).nodeValue
                Case "Ct:3" '18
                    gCobas4800Res.Type18_CT = xNode.Attributes(1).nodeValue
                Case "Result 1" 'Other
                    If InStr(1, xNode.Attributes(1).nodeValue, "POS") > 0 Then
                        gCobas4800Res.TypeOther_Res = "positive"
                    ElseIf InStr(1, xNode.Attributes(1).nodeValue, "NEG") > 0 Then
                        gCobas4800Res.TypeOther_Res = "negative"
                    End If
                    
                Case "Result 2" '16
                    If InStr(1, xNode.Attributes(1).nodeValue, "POS") > 0 Then
                        gCobas4800Res.Type16_Res = "positive"
                    ElseIf InStr(1, xNode.Attributes(1).nodeValue, "NEG") > 0 Then
                        gCobas4800Res.Type16_Res = "negative"
                    End If
                Case "Result 3" '18
                    If InStr(1, xNode.Attributes(1).nodeValue, "POS") > 0 Then
                        gCobas4800Res.Type18_Res = "positive"
                    ElseIf InStr(1, xNode.Attributes(1).nodeValue, "NEG") > 0 Then
                        gCobas4800Res.Type18_Res = "negative"
                    End If
                
                
                End Select

                
            End Select

        
        End If
        
        If xNode.hasChildNodes Then
            Cobas4800_Xml_DisPlay xNode.childNodes, Indent
        End If
    Next xNode
    
End Sub


