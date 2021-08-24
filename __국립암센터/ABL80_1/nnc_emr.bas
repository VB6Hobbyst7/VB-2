Attribute VB_Name = "nnc_emr"
Option Explicit

Public gOnline_Test As String
Public gServerPath As String
Public giIndex  As Long
Public gOrderExam As String

'Public Const gXml_S03 = "PG_SRL.SLP91_S03"
'Public Const gXml_S07 = "PG_SRL.SLP91_S07"

Public Const gXml_S01 = "PG_SRL.SLP91_S01"
Public Const gXml_S02 = "PG_SRL.SLP91_S02"
Public Const gXml_S03 = "PG_SRL.SLP91_S03"

Public Const gXml_S04 = "PG_SRL.SLP91_S04"
Public Const gXml_S07 = "PG_SRL.SLP91_S07"
Public Const gXml_S18 = "PG_SRL.SLP91_S18"

Public Const gXml_S10 = "PG_SRL.SLP91_S10"
Public Const gXml_S11 = "PG_SRL.SLP91_S11"
Public Const gXml_S13 = "PG_SRL.SLP91_S13"

Public Const gXml_S26 = "PG_SRL.SLP91_S26"
Public Const gXml_S27 = "PG_SRL.SLP91_S27"

Public Const gXml_U01 = "PG_SRL.SLP91_U01"
Public Const gXml_U03 = "PG_SRL.SLP91_U03"
Public Const gXml_U06 = "PG_SRL.SLP91_U06"


Type Exam_Select
    TST_CD      As String
    TST_CNT     As Integer
End Type


Type QC_Select
    ORDDATE     As String
    QMCODE      As String
    LOTNO       As String
    ORDSEQNO    As String
    EQIPCODE    As String
    EXAMCODE    As String
    ROOMCODE    As String
    EXAMNAME    As String
    ORDYN       As String
    SETEXYN     As String
    SPCCODE     As String
    SPCNAME     As String
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
    ORD_SITE    As String
    TST_CLS     As String
    RERUN       As String
    TST_FRCT_CD1    As String
    HSP_CLS     As String
    ACPT_DTETM  As String
    AGE         As String
    OK          As Integer
End Type

Type Doctor_Select
    WKPERS_ID   As String
    WKPERS_NM   As String
End Type

Public gQC_Select() As QC_Select
Public gExam_Select() As Exam_Select
Public gPat_Info_Select As PatInfo_Select
Public gQC_Rece() As String
Public gDoctor() As Doctor_Select

Public gEMRBarcode As String


Public Sub Clear_Doctor()
    giIndex = -1
    ReDim gDoctor(0)
    
End Sub

Public Sub Clear_QC_Rece()
    giIndex = -1
    ReDim gQC_Rece(0)
    
End Sub


Public Sub Clear_XML_Exam()
    giIndex = -1
    ReDim gExam_Select(0)
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

Public Sub Clear_XML_QC()
    giIndex = -1
    ReDim gQC_Select(0)
End Sub

Public Function Online_Param(ByVal asProc As String, ByVal asParam As String) As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    Online_Param = ""
    sFileName = "Res"
    
    sParam = asParam
    
    sRetStr = Online_XML_Qry(asProc, sParam)
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    Dim xDoc As MSXML.DOMDocument
    Set xDoc = New MSXML.DOMDocument
    If xDoc.Load(App.Path & "\XML\" & sFileName & ".xml") Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_U03
            Clear_QC_Rece
            display_online_parsing_QCRece xDoc.childNodes, 0
        Case gXml_U06
            gEMRBarcode = ""
            
            display_online_parsing_SRece xDoc.childNodes, 0
        End Select
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
        Online_Param = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If
    
End Function


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
        Case gXml_S11
            Clear_XML_QC
            display_online_parsing_QC xDoc.childNodes, 0
        
            
'        Case gXml_S04
''            Clear_XML_TLA
'            display_online_parsing_TLAInfo xDoc.childNodes, 0
'        Case gXml_S13
'            gS13_WorkList_Clear
'            display_online_parsing_S13 xDoc.childNodes, 0
            
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

Private Function TLA_Param(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String
    Dim sProc As String
    Dim sParam As String
    
    TLA_Param = ""
    sProc = asProc
    
    Select Case sProc
    Case gXml_S04, gXml_S13, gXml_S11
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
                 
    Case gXml_U03
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


Public Function Online_XML(ByVal asProc As String, ByVal asSpcno As String) As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    Online_XML = ""
    sFileName = "Res"
    
    sParam = Select_Param(asProc, asSpcno)
    
    sRetStr = Online_XML_Qry(asProc, sParam)
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    Dim xDoc As MSXML.DOMDocument
    Set xDoc = New MSXML.DOMDocument
    If xDoc.Load(App.Path & "\XML\" & sFileName & ".xml") Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_S03
            Clear_XML_PInfo
            display_online_parsing_PatInfo xDoc.childNodes, 0
        Case gXml_S07
            Clear_XML_Exam
            display_online_parsing_ExamCode xDoc.childNodes, 0
        Case gXml_S26
            Clear_Doctor
            display_online_parsing_Dr xDoc.childNodes, 0
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
    
    Case gXml_S03
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>"
    
    Case gXml_S07
    
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>"
    Case gXml_S26
    
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>"
                 
'        <Table>
'<QID><![CDATA[PG_SRL.SLP91_S26]]></QID>
'<QTYPE><![CDATA[Package]]></QTYPE>
'<USERID><![CDATA[LIA]]></USERID>
'<EXECTYPE><![CDATA[FILL]]></EXECTYPE>
'<TABLENAME><![CDATA[]]></TABLENAME>
'<P0><![CDATA[SVAN]]></P0>
'<P1><![CDATA[]]></P1></Table>
        
                 
    End Select
    
    Select_Param = sParam
    
End Function

'XML File Parsing===========================================================================================================
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

Public Sub display_online_parsing_Dr(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)

    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            Select Case xNode.parentNode.nodeName
            Case "WKPERS_ID"
                giIndex = giIndex + 1
                ReDim Preserve gDoctor(giIndex)
                gDoctor(giIndex).WKPERS_ID = xNode.nodeValue
            Case "WKPERS_NM"
                gDoctor(giIndex).WKPERS_NM = xNode.nodeValue


            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_Dr xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub display_online_parsing_SRece(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)

    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            If Len(Trim(xNode.nodeValue)) = 11 Then
                gEMRBarcode = xNode.nodeValue
            End If
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_SRece xNode.childNodes, Indent
        End If
        
    Next xNode
End Sub

Public Sub display_online_parsing_QCRece(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)

    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
    
            giIndex = giIndex + 1
            ReDim Preserve gQC_Rece(giIndex)
            gQC_Rece(giIndex) = xNode.nodeValue
            gQC_Rece(giIndex) = Replace(gQC_Rece(giIndex), "〓", "")
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_QCRece xNode.childNodes, Indent
        End If
    Next xNode
End Sub


Public Sub display_online_parsing_QC(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)

    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            Select Case xNode.parentNode.nodeName
            Case "ORDDATE"
                giIndex = giIndex + 1
                ReDim Preserve gQC_Select(giIndex)
                gQC_Select(giIndex).ORDDATE = xNode.nodeValue
            Case "QMCODE"
                gQC_Select(giIndex).QMCODE = xNode.nodeValue
            Case "LOTNO"
                gQC_Select(giIndex).LOTNO = xNode.nodeValue
            Case "ORDSEQNO"
                gQC_Select(giIndex).ORDSEQNO = xNode.nodeValue
            Case "EQIPCODE"
                gQC_Select(giIndex).EQIPCODE = xNode.nodeValue
            Case "EXAMCODE"
                gQC_Select(giIndex).EXAMCODE = xNode.nodeValue
            Case "ROOMCODE"
                gQC_Select(giIndex).ROOMCODE = xNode.nodeValue
            Case "EXAMNAME"
                gQC_Select(giIndex).EXAMNAME = xNode.nodeValue
            Case "ORDYN"
                gQC_Select(giIndex).ORDYN = xNode.nodeValue
            Case "SETEXYN"
                gQC_Select(giIndex).SETEXYN = xNode.nodeValue
            Case "SPCCODE"
                gQC_Select(giIndex).SPCCODE = xNode.nodeValue
            Case "SPCNAME"
                gQC_Select(giIndex).SPCNAME = xNode.nodeValue

            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_QC xNode.childNodes, Indent
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
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_PatInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub Display_Online_Parsing(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
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
Public Function Online_Result(ByVal asParam As String) As String

    Dim sRetStr As String


    Online_Result = ""

    gOnline_Ret = ""

    sRetStr = Online_Result_Qry(asParam)

    'SaveXMLFile sRetStr
    Xml_Log sRetStr, "res"

    Dim xDoc As MSXML.DOMDocument

    Set xDoc = New MSXML.DOMDocument

    If xDoc.Load(App.Path & "\Res\res.xml") Then
    'If xDoc.Load(sRetStr) Then
        ' 문서가 성공적으로 로드되었습니다.
        ' 이제 재미있는 작업을 수행합니다.
        Display_Online_Parsing xDoc.childNodes, 0
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
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        SaveXML_Data Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        SaveXML_Data Format(Time, "hh:nn:ss") & "[SOAP]" & send
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
End Function

Public Function Online_Result_Qry_Conf(ByVal asParam As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String

    On Error GoTo ErrHandle

    Set oSOAP = New MSSOAPLib30.SoapClient30

    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit gServerPath
    
    strDiv = "PG_SRL.SLP91_U07"
    
    sParam = asParam
    
    SaveXML_Data "[Save Result]" & sParam
    send = oSOAP.wsLISInterface(strDiv, sParam)
    SaveXML_Data "[Save Result => Return]" & send
    Online_Result_Qry_Conf = send
    Set oSOAP = Nothing
    DoEvents
    Exit Function

ErrHandle:
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        SaveXML_Data Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        SaveXML_Data Format(Time, "hh:nn:ss") & "[SOAP]" & send
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
End Function

Public Function Online_Result_Qry_Conf_Cancel(ByVal asParam As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String

    On Error GoTo ErrHandle

    Set oSOAP = New MSSOAPLib30.SoapClient30

    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit gServerPath
    
    strDiv = "PG_SRL.SLP91_U08"
    
    sParam = asParam
    
    SaveXML_Data "[Save Result]" & sParam
    send = oSOAP.wsLISInterface(strDiv, sParam)
    SaveXML_Data "[Save Result => Return]" & send
    Online_Result_Qry_Conf_Cancel = send
    Set oSOAP = Nothing
    DoEvents
    Exit Function

ErrHandle:
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        SaveXML_Data Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        SaveXML_Data Format(Time, "hh:nn:ss") & "[SOAP]" & send
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
End Function

'==================================================================================================== Result Trans sub start



''Result Trans sub start ====================================================================================================
'Public Function Online_Result_New(ByVal asSpcno As String, _
'                              ByVal asExam As String, _
'                              ByVal asRes As String, _
'                              ByVal asEquip As String, _
'                              ByVal asCount As String, _
'                              ByVal asEqFlag As String, _
'                              ByVal asUser As String) As String
'
'    Dim sRetStr As String
'
'
'    Online_Result_New = ""
'
'    gOnline_Ret = ""
'
'    sRetStr = Online_Result_Qry_New(asSpcno, asExam, asRes, asEquip, asCount, asEqFlag, asUser)
'
'    'SaveXMLFile sRetStr
'    Save_Xml_Data sRetStr, "res"
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
'        Online_Result_New = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
'    End If
'
'End Function
'
'Public Function Online_Result_Qry_New(ByVal asSpcno As String, _
'                              ByVal asExam As String, _
'                              ByVal asRes As String, _
'                              ByVal asEquip As String, _
'                              ByVal asCount As String, _
'                              ByVal asEqFlag As String, _
'                              ByVal asUser As String) As String
'    Dim oSOAP As MSSOAPLib30.SoapClient30
'    Dim strDiv As String
'    Dim send As String
'    Dim sParam As String
'
'    On Error GoTo ErrHandle
'
'    Set oSOAP = New MSSOAPLib30.SoapClient30
'
'    oSOAP.ClientProperty("ServerHTTPRequest") = True
'
'    oSOAP.MSSoapInit gServerPath
'
'    strDiv = "PG_SRL.SLP91_P03"
'
'    sParam = "<Table>" & _
'             "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
'             "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'             "<USERID><![CDATA[LIA]]></USERID>" & _
'             "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'             "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'             "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
'             "<P1><![CDATA[" & asExam & "]]></P1>" & _
'             "<P2><![CDATA[" & asRes & "]]></P2>" & _
'             "<P3><![CDATA[" & asEqFlag & "]]></P3>" & _
'             "<P4><![CDATA[" & asEquip & "]]></P4>" & _
'             "<P5><![CDATA[]]></P5>" & _
'             "<P6><![CDATA[" & asCount & "]]></P6>" & _
'             "<P7><![CDATA[]]></P7>" & _
'             "<P8><![CDATA[]]></P8>" & _
'             "<P9><![CDATA[" & asUser & "]]></P9>" & _
'             "</Table>"
'    SaveXML_Data "[Save Result]" & sParam
'    send = oSOAP.wsLISInterface(strDiv, sParam)
'    SaveXML_Data "[Save Result => Return]" & send
'    Online_Result_Qry_New = send
'    Set oSOAP = Nothing
'    DoEvents
'    Exit Function
'
'ErrHandle:
'    If oSOAP.FaultString <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
'    End If
'    If Trim(Err.Description) <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
'    End If
'End Function
'
''==================================================================================================== Result Trans sub start



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

