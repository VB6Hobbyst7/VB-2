Attribute VB_Name = "modNCCEmr"
Option Explicit

Public gOnline_Test As String
Public gServerPath  As String
Public gIFUser      As String
Public giIndex      As Long
Public gPanel       As String
Public gDoctor      As String
'gOrderExam

Public Const gXml_S01 = "PG_SRL.SLP91_S01"
Public Const gXml_S02 = "PG_SRL.SLP91_S02"
Public Const gXml_S03 = "PG_SRL.SLP91_S03"
Public Const gXml_S04 = "PG_SRL.SLP91_S04"
Public Const gXml_S07 = "PG_SRL.SLP91_S07"
Public Const gXml_S10 = "PG_SRL.SLP91_S10"
Public Const gXml_S18 = "PG_SRL.SLP91_S18"
Public Const gXml_U07 = "PG_SRL.SLP91_U01"

'-- POCT
Public Const gXml_S28 = "PG_SRL.SLP91_S28"
Public Const gXml_U06 = "PG_SRL.SLP91_U06"
Public Const gXml_P03 = "PG_SRL.SLP91_P03"


Type Exam_Select
    TST_CD      As String
    TST_CNT     As Integer
End Type

Type PatInfo_Select
    TST_CD          As String
    TST_NM          As String
    TST_FRCT_CD     As String
    ACPTNO_1        As String
    PT_NO           As String
    PT_NM           As String
    Sex             As String
    SPC_CD_1        As String
    ORD_SITE        As String
    TST_CLS         As String
    RERUN           As String
    TST_FRCT_CD1    As String
    HSP_CLS         As String
    ACPT_DTETM      As String
    Age             As String
    SPC_NO          As String
        
    MEDDEPT         As String   '진료과
    GENDR           As String   '담당의
        
    BARCODE         As String   '검체번호
    RCPNO           As String   '접수번호
    OK              As Integer
End Type


Public gExam_Select() As Exam_Select
Public gPat_Info_Select As PatInfo_Select


Public Sub Clear_XML_Exam()
    giIndex = -1
    ReDim gExam_Select(0)
End Sub

Public Sub Clear_XML_PInfo()
    gPat_Info_Select.ACPT_DTETM = ""
    gPat_Info_Select.ACPTNO_1 = ""
    gPat_Info_Select.Age = ""
    gPat_Info_Select.HSP_CLS = ""
    gPat_Info_Select.OK = -1
    gPat_Info_Select.ORD_SITE = ""
    gPat_Info_Select.PT_NM = ""
    gPat_Info_Select.PT_NO = ""
    gPat_Info_Select.RERUN = ""
    gPat_Info_Select.Sex = ""
    gPat_Info_Select.SPC_CD_1 = ""
    gPat_Info_Select.TST_CD = ""
    gPat_Info_Select.TST_CLS = ""
    gPat_Info_Select.TST_FRCT_CD = ""
    gPat_Info_Select.TST_FRCT_CD1 = ""
    gPat_Info_Select.TST_NM = ""
    gPat_Info_Select.SPC_NO = ""
    gPat_Info_Select.ACPT_DTETM = ""
    
    gPat_Info_Select.MEDDEPT = ""
    gPat_Info_Select.GENDR = ""
    gPat_Info_Select.BARCODE = ""
    gPat_Info_Select.RCPNO = ""
    
End Sub

'Public Function Online_TLA(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String
'
'    Dim sRetStr As String
'    Dim sFileName As String
'    Dim sParam As String
'
'    Online_TLA = ""
''    gOrderExam = ""
'    sFileName = "Res"
'
'    sParam = TLA_Param(asProc, asDate1, asDate2, asBarcode)
'
'    sRetStr = Online_XML_Qry(asProc, sParam)
'
'    'SaveXMLFile sRetStr
'    Xml_Log sRetStr, sFileName
'
'    Dim xDoc As MSXML2.DOMDocument
'    Set xDoc = New MSXML2.DOMDocument
'    If xDoc.Load(App.Path & "\XML\" & sFileName & ".xml") Then
'        ' Data Load, Start Parsing
'        Select Case asProc
'        Case gXml_S18
'            Clear_XML_PInfo
'
'            display_online_parsing_TLAInfo xDoc.childNodes, 0
'        End Select
''        Display_Online_Parsing_Test xDoc.childNodes, 0
'    Else
'        ' 문서를 로드하지 못했습니다.
'        Dim strErrText As String
'        Dim xPE As MSXML2.IXMLDOMParseError
'       ' ParseError 개체를 가져옵니다
'        Set xPE = xDoc.parseError
'        With xPE
'
'            strErrText = "Your XML Document failed to load" & _
'                         "due the following error." & vbCrLf & _
'                         "Error #: " & .errorCode & ": " & xPE.reason & _
'                         "Line #: " & .Line & vbCrLf & _
'                         "Line Position: " & .linepos & vbCrLf & _
'                         "Position In File: " & .filepos & vbCrLf & _
'                         "Source Text: " & .srcText & vbCrLf & _
'                         "Document URL: " & .URL
'        End With
'
''        SaveXML_Data strErrText
'    End If
'
'    Xml_Log gOnline_Test, "TLA"
'
'
'    Set xPE = Nothing
'    Set xDoc = Nothing
'
'    If InStr(1, gOnline_Ret, vbTab) > 0 Then
'        Online_TLA = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
'    End If
'
'End Function

Public Function Online_XML(ByVal asProc As String, ByVal asSpcno As String, Optional ByVal asDate As String) As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    Online_XML = ""
    sFileName = "Res"
    
    sParam = Select_Param(asProc, asSpcno, asDate)
    Call SetSQLData("sParam", sParam)
    
    sRetStr = Online_XML_Qry(asProc, sParam)
    Call SetSQLData("sRetStr", sRetStr)
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    Dim xDoc As MSXML2.DOMDocument
    Set xDoc = New MSXML2.DOMDocument
    
    If xDoc.Load(App.Path & "\XML\" & sFileName & ".xml") Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_S03
            Clear_XML_PInfo
            display_online_parsing_PatInfo xDoc.childNodes, 0
        Case gXml_S07
            Clear_XML_Exam
            display_online_parsing_ExamCode xDoc.childNodes, 0
        Case gXml_S28
            Clear_XML_PInfo
            display_online_parsing_doctInfo xDoc.childNodes, 0
        Case gXml_U06
            'Clear_XML_PInfo
            display_online_parsing_poctInfo xDoc.childNodes, 0
            
        End Select
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
    Case gXml_S07
        If gExam_Select(0).TST_CNT > 0 Then
            Online_XML = 1
        Else
            Online_XML = 0
        End If
    
    Case gXml_S28
        If gPat_Info_Select.OK = 1 Then
            Online_XML = 1
        Else
            Online_XML = 0
        End If
    
    Case gXml_U06
        If gPat_Info_Select.OK = 1 Then
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

    SaveXML_Data "[Use Proc => " & strDiv & " ]" & sParam
    send = oSOAP.wsLISInterface(strDiv, sParam)
    SaveXML_Data "[Return Proc => " & strDiv & " ]" & send
    Online_XML_Qry = send
    Set oSOAP = Nothing
    DoEvents
    Exit Function
    
ErrHandle:
    Call SetSQLData("XMLQRY", Err.Number & Err.Description & vbCrLf & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf)
    
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        'Call SetSQLData("oSOAP.FaultString :", Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf)
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
        'Call SetSQLData("Err.Description : ", Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf)
    End If
    
End Function

Private Function Select_Param(ByVal asProc As String, ByVal asSpcno As String, Optional ByVal asDate As String) As String
    Dim sProc   As String
    Dim sParam  As String
    Dim strDate As String
    Dim strRcpR As String   '발행처
    'Dim strRcpM As String   '발행자
    'Dim strItem As String
    
    Select_Param = ""
    sProc = asProc
    
    If sProc = gXml_U06 Then
        '실제 검사일시
        strDate = asDate
    Else
        strDate = Format(Now, "yyyymmddhhmmss")
    End If
    
    strRcpR = ""
    'strRcpM = gDoctor
    'strItem = gPanel
    
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
    Case gXml_S03
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[SUADEV]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>"
    Case gXml_S07
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[SUADEV]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>"
    
    Case gXml_U06
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[SUA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & gEquipCode & "]]></P0>" & _
                 "<P1><![CDATA[" & asSpcno & "]]></P1>" & _
                 "<P2><![CDATA[" & strDate & "]]></P2>" & _
                 "<P3><![CDATA[" & strRcpR & "]]></P3>" & _
                 "<P4><![CDATA[" & gDoctor & "]]></P4>" & _
                 "<P5><![CDATA[" & gPanel & "]]></P5>" & _
                 "<P6><![CDATA[" & gIFUser & "]]></P6>" & _
                 "<P7><![CDATA[]]></P7>" & _
                 "<P8><![CDATA[]]></P8>" & _
                 "<P9><![CDATA[]]></P9>" & _
                 "<P10><![CDATA[]]></P10>" & _
                 "</Table>"
    
'        sParam = "<?xml version='1.0' encoding='UTF-8'?>" & _
                 "<NewDataSet>" & _
                 "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[SUA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & gEquipCode & "]]></P0>" & _
                 "<P1><![CDATA[" & asSpcno & "]]></P1>" & _
                 "<P2><![CDATA[" & strDate & "]]></P2>" & _
                 "<P3><![CDATA[" & strRcpR & "]]></P3>" & _
                 "<P4><![CDATA[" & gDoctor & "]]></P4>" & _
                 "<P5><![CDATA[" & gPanel & "]]></P5>" & _
                 "<P6><![CDATA[" & gIFUser & "]]></P6>" & _
                 "<P7><![CDATA[]]></P7>" & _
                 "<P8><![CDATA[]]></P8>" & _
                 "<P9><![CDATA[]]></P9>" & _
                 "<P10><![CDATA[]]></P10>" & _
                 "</Table>" & _
                 "</NewDataSet>"
    
    Case gXml_S28
        sParam = "<Table>" & _
                 "<QID><![CDATA[PG_SRL.SLP91_S28]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[SUA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>"
                 
'        sParam = "<?xml version='1.0' encoding='UTF-8'?>" & _
                 "<NewDataSet>" & _
                 "<Table>" & _
                 "<QID><![CDATA[PG_SRL.SLP91_S28]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[SUA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>" & _
                 "</NewDataSet>"
                 

    End Select
    
    Select_Param = sParam
    
End Function

Private Function TLA_Param(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String
    Dim sProc As String
    Dim sParam As String
    
    TLA_Param = ""
    sProc = asProc
    
    Select Case sProc
    Case gXml_S04
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
    Case gXml_S18
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asDate1 & "]]></P0>" & _
                 "<P1><![CDATA[" & asDate2 & "]]></P1>" & _
                 "<P2><![CDATA[" & asBarcode & "]]></P2>" & _
                 "<P3><![CDATA[]]></P3>" & _
                 "</Table>"
    End Select
    
    TLA_Param = sParam
    
End Function


'XML File Parsing===========================================================================================================
Public Sub display_online_parsing_ExamCode(ByRef Nodes As MSXML2.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
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

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            Select Case xNode.parentNode.nodeName
            Case "TST_CD"
                gPat_Info_Select.TST_CD = xNode.nodeValue
                gPat_Info_Select.OK = 1
            Case "ACPT_DTETM"
                gPat_Info_Select.ACPT_DTETM = xNode.nodeValue
            Case "ACPTNO_1"
                gPat_Info_Select.ACPTNO_1 = xNode.nodeValue
            Case "AGE"
                gPat_Info_Select.Age = xNode.nodeValue
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
                gPat_Info_Select.Sex = xNode.nodeValue
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
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_PatInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub display_online_parsing_poctInfo(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2
    
    'io_spc_no '검체
    'io_acptno '접수

    For Each xNode In Nodes
        If xNode.nodeType = 4 Then
        
'Call SetSQLData("NODE", xNode.parentNode.nodeName & "")
'Call SetSQLData("NODE", xNode.nodeValue & "")

            Select Case xNode.parentNode.nodeName
            Case "Value"    '검체번호
                If Len(xNode.nodeValue & "") = 11 Then
                    gPat_Info_Select.BARCODE = xNode.nodeValue
                End If
            'Case "IO_ACPTNO"    '접수번호
            '    gPat_Info_Select.RCPNO = xNode.nodeValue
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_poctInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub display_online_parsing_doctInfo(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
        If xNode.nodeType = 4 Then
            
'Call SetSQLData("NODE", xNode.parentNode.nodeName & "")
'Call SetSQLData("NODE", xNode.nodeValue & "")

            Select Case xNode.parentNode.nodeName
            Case "MEDDEPT"
                gPat_Info_Select.MEDDEPT = xNode.nodeValue
            Case "GENDR"
                gPat_Info_Select.GENDR = xNode.nodeValue
                gDoctor = gPat_Info_Select.GENDR
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_doctInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub display_online_parsing_TLAInfo(ByRef Nodes As MSXML2.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "TST_CD"
                gPat_Info_Select.TST_CD = xNode.nodeValue
                If gOrderExam = "" Then
                    gOrderExam = "'" & xNode.nodeValue & "'"
                Else
                    gOrderExam = gOrderExam & ", '" & xNode.nodeValue & "'"
                End If
                gPat_Info_Select.OK = 1
            Case "ACPT_DTETM"
                gPat_Info_Select.ACPT_DTETM = xNode.nodeValue
            Case "ACPTNO_1"
                gPat_Info_Select.ACPTNO_1 = xNode.nodeValue
            Case "AGE"
                gPat_Info_Select.Age = xNode.nodeValue
'                Exit Sub
            Case "SPC_NO"
                gPat_Info_Select.SPC_NO = xNode.nodeValue
            Case "ORD_SITE"
                gPat_Info_Select.ORD_SITE = xNode.nodeValue
            Case "PT_NM"
                gPat_Info_Select.PT_NM = xNode.nodeValue
            Case "PT_NO"
                gPat_Info_Select.PT_NO = xNode.nodeValue
            Case "SEX"
                gPat_Info_Select.Sex = xNode.nodeValue
            Case "SPC_CD_1"
                gPat_Info_Select.SPC_CD_1 = xNode.nodeValue
            Case "TST_CLS"
                gPat_Info_Select.TST_CLS = xNode.nodeValue
            Case "TST_FRCT_CD"
                gPat_Info_Select.TST_FRCT_CD = xNode.nodeValue
            Case "TST_NM"
                gPat_Info_Select.TST_NM = xNode.nodeValue
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_TLAInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub Display_Online_Parsing(ByRef Nodes As MSXML2.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab

        End If
        If xNode.hasChildNodes Then
            Display_Online_Parsing xNode.childNodes, Indent
        End If
    Next xNode
End Sub

'===========================================================================================================XML File Parsing

'Result Trans sub start ====================================================================================================
'Public Function Online_Result(ByVal asParam As String) As String
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
'    Dim xDoc As MSXML2.DOMDocument
'
'    Set xDoc = New MSXML2.DOMDocument
'
'    If xDoc.Load(App.Path & "\Res\res.xml") Then
'    'If xDoc.Load(sRetStr) Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_Online_Parsing xDoc.childNodes, 0
'    Else
'        ' 문서를 로드하지 못했습니다.
'        Dim strErrText As String
'        Dim xPE As MSXML2.IXMLDOMParseError
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
'                         "Document URL: " & .URL
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
'End Function

Public Function Online_Result_Qry(ByVal asParam As String, ByVal strDiv As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    'Dim strDiv As String
    Dim send As String
    Dim sParam As String

    On Error GoTo ErrHandle

    Set oSOAP = New MSSOAPLib30.SoapClient30

    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit gServerPath
    
'    strDiv = "PG_SRL.SLP91_P03"
    
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



