Attribute VB_Name = "nnc_emr"
Option Explicit

Public gOnline_Test As String
Public gServerPath As String
Public giIndex  As Long

Public Const gXml_S01 = "PG_SRL.SLP91_S01"
Public Const gXml_S02 = "PG_SRL.SLP91_S02"
Public Const gXml_S03 = "PG_SRL.SLP91_S03"      '환자정보,검사정보

Public Const gXml_S04 = "PG_SRL.SLP91_S04"
Public Const gXml_S07 = "PG_SRL.SLP91_S07"      '검사정보

Public Const gXml_S10 = "PG_SRL.SLP91_S10"

Public Const gXml_S18 = "PG_SRL.SLP91_S18"      '검사일자,접수번호,검사파트

Public Const gXml_U07 = "PG_SRL.SLP91_U01"

Type Exam_Select
    TST_CD      As String
    TST_CNT     As Integer
    OK          As Integer
End Type
Public gExam_Select() As Exam_Select

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
Public gPat_Info_Select As PatInfo_Select

Type TLAInfo_Select
    TST_DTE     As String
    SPCNO       As String
    OK          As Integer
End Type
Public gTLA_Info_Select() As TLAInfo_Select

Type Urine_Select
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
    SPC_NO      As String
    OK          As Integer
End Type
Public gUrine_Info_Select() As Urine_Select

Public Sub Clear_XML_Exam()
    giIndex = -1
    ReDim gExam_Select(0)
End Sub

Public Sub Clear_XML_TLAInfo()
    giIndex = -1
    ReDim gTLA_Info_Select(0)
End Sub

Public Sub Clear_XML_UrineInfo()
    giIndex = -1
    ReDim gUrine_Info_Select(0)
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

Public Function Online_Urine(ByVal asProc As String, ByVal asDate As String, ByVal asAcpNo As String, ByVal asPart As String) As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    Online_Urine = ""
    sFileName = "Res"
    
    sParam = Urine_Param(asProc, asDate, asAcpNo, asPart)
    
    sRetStr = Online_XML_Qry(asProc, sParam)
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    Dim xDoc As MSXML.DOMDocument
    Set xDoc = New MSXML.DOMDocument
    If xDoc.Load(App.Path & "\XML\" & sFileName & ".xml") Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_S18
            Clear_XML_UrineInfo
            display_online_parsing_UrineInfo xDoc.childNodes, 0
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
    
    If sRetStr = "" Then
        Online_Urine = 0
    Else
        If gUrine_Info_Select(giIndex).OK = 1 Then
            Online_Urine = 1
        Else
            Online_Urine = 0
        End If
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
    
'    If InStr(1, gOnline_Ret, vbTab) > 0 Then
'        Online_XML = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
'    End If
    
    Select Case asProc
    Case gXml_S03
        If gPat_Info_Select.OK = 1 Then
            Online_XML = 1
        Else
            Online_XML = 0
        End If
    Case gXml_S07
        If gExam_Select(0).OK = 1 Then
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
    Case gXml_S01, gXml_S02, gXml_S03, gXml_S07, gXml_S10
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[" & gServerID & "]]></USERID>" & _
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
    Case gXml_S04
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[" & gServerID & "]]></USERID>" & _
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
                 "<USERID><![CDATA[" & gServerID & "]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asBarcode & "]]></P0>" & _
                 "<P1><![CDATA[" & asDate1 & "]]></P1>" & _
                 "<P2><![CDATA[" & asDate2 & "]]></P2>" & _
                 "</Table>"
    End Select
    
    TLA_Param = sParam
    
End Function

Private Function Urine_Param(ByVal asProc As String, ByVal asDate As String, ByVal asAcpNo As String, ByVal asPart As String) As String
    Dim sProc As String
    Dim sParam As String
    
    Urine_Param = ""
    sProc = asProc
    
    Select Case sProc
    Case gXml_S18
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[" & gServerID & "]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asDate & "]]></P0>" & _
                 "<P1><![CDATA[" & asAcpNo & "]]></P1>" & _
                 "<P2><![CDATA[" & asPart & "]]></P2>" & _
                 "<P3><![CDATA[]]></P3>" & _
                 "</Table>"
    End Select
    
    Urine_Param = sParam
    
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
                gExam_Select(giIndex).TST_CNT = giIndex + 1
            End Select
            gExam_Select(0).OK = 1
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

Public Sub display_online_parsing_UrineInfo(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "TST_CD"
                giIndex = giIndex + 1
                ReDim Preserve gUrine_Info_Select(giIndex)
                
                gUrine_Info_Select(giIndex).TST_CD = xNode.nodeValue
            Case "TST_NM"
                gUrine_Info_Select(giIndex).TST_NM = xNode.nodeValue
            Case "TST_FRCT_CD"
                gUrine_Info_Select(giIndex).TST_FRCT_CD = xNode.nodeValue
            Case "ACPTNO_1"
                gUrine_Info_Select(giIndex).ACPTNO_1 = xNode.nodeValue
            Case "PT_NO"
                gUrine_Info_Select(giIndex).PT_NO = xNode.nodeValue
            Case "PT_NM"
                gUrine_Info_Select(giIndex).PT_NM = xNode.nodeValue
            Case "SEX"
                gUrine_Info_Select(giIndex).SEX = xNode.nodeValue
            Case "SPC_CD_1"
                gUrine_Info_Select(giIndex).SPC_CD_1 = xNode.nodeValue
            Case "ORD_SITE"
                gUrine_Info_Select(giIndex).ORD_SITE = xNode.nodeValue
            Case "TST_CLS"
                gUrine_Info_Select(giIndex).TST_CLS = xNode.nodeValue
            Case "RERUN"
                gUrine_Info_Select(giIndex).RERUN = xNode.nodeValue
            Case "TST_FRCT_CD1"
                gUrine_Info_Select(giIndex).TST_FRCT_CD1 = xNode.nodeValue
            Case "HSP_CLS"
                gUrine_Info_Select(giIndex).HSP_CLS = xNode.nodeValue
            Case "ACPT_DTETM"
                gUrine_Info_Select(giIndex).ACPT_DTETM = xNode.nodeValue
            Case "AGE"
                gUrine_Info_Select(giIndex).AGE = xNode.nodeValue
            Case "SPC_NO"
                gUrine_Info_Select(giIndex).SPC_NO = xNode.nodeValue
            End Select
            gUrine_Info_Select(giIndex).OK = 1
            
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_UrineInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

'===========================================================================================================XML File Parsing

'Result Trans sub start ====================================================================================================
Public Function Online_Result(ByVal asParam As String) As String

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

