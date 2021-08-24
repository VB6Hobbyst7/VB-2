Attribute VB_Name = "nnc_emr1"
Option Explicit

Public gOnline_Ret1 As String
Public gOnline_Test1 As String
Public giIndex1  As Long

Type Exam_Select
    TST_CD      As String
    TST_CNT     As Integer
    OK          As Integer
End Type
Public gExam_Select1() As Exam_Select

Type PatInfo_Select
    TST_CD      As String
    TST_NM      As String
    TST_FRCT_CD As String
    ACPTNO_1    As String
    PT_NO       As String
    PT_NM       As String
    Sex         As String
    SPC_CD_1    As String
    ORD_SITE    As String
    TST_CLS     As String
    RERUN       As String
    TST_FRCT_CD1    As String
    HSP_CLS     As String
    ACPT_DTETM  As String
    Age         As String
    OK          As Integer
End Type
Public gPat_Info_Select1 As PatInfo_Select

Type TLAInfo_Select
    TST_DTE     As String
    SPCNO       As String
    OK          As Integer
End Type
Public gTLA_Info_Select1() As TLAInfo_Select

Public Sub Clear_XML_TLALIST()
    giIndex1 = -1
    ReDim gTLA_Info_Select(0)
End Sub

Public Sub Clear_XML_Exam1()
    giIndex1 = -1
    ReDim gExam_Select1(0)
End Sub

Public Sub Clear_XML_PInfo1()
    gPat_Info_Select1.ACPT_DTETM = ""
    gPat_Info_Select1.ACPTNO_1 = ""
    gPat_Info_Select1.Age = ""
    gPat_Info_Select1.HSP_CLS = ""
    gPat_Info_Select1.OK = -1
    gPat_Info_Select1.ORD_SITE = ""
    gPat_Info_Select1.PT_NM = ""
    gPat_Info_Select1.PT_NO = ""
    gPat_Info_Select1.RERUN = ""
    gPat_Info_Select1.Sex = ""
    gPat_Info_Select1.SPC_CD_1 = ""
    gPat_Info_Select1.TST_CD = ""
    gPat_Info_Select1.TST_CLS = ""
    gPat_Info_Select1.TST_FRCT_CD = ""
    gPat_Info_Select1.TST_FRCT_CD1 = ""
    gPat_Info_Select1.TST_NM = ""
    
End Sub

Public Function Online_TLA1(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    Online_TLA1 = ""
    sFileName = "Res"
    
    sParam = TLA_Param1(asProc, asDate1, asDate2, asBarcode)
    
    sRetStr = Online_XML_Qry1(asProc, sParam)
    
    'SaveXMLFile sRetStr
    Xml_Log1 sRetStr, sFileName
    
    Dim xDoc As MSXML.DOMDocument
    Set xDoc = New MSXML.DOMDocument
    If xDoc.Load(App.Path & "\XML\" & sFileName & ".xml") Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_S04
'            Clear_XML_TLA
            display_online_parsing_TLAInfo1 xDoc.childNodes, 0
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
    
    Xml_Log1 gOnline_Test, "TLA"
    
    
    Set xPE = Nothing
    Set xDoc = Nothing
    
    If InStr(1, gOnline_Ret1, vbTab) > 0 Then
        Online_TLA1 = Left(gOnline_Ret1, InStr(1, gOnline_Ret1, vbTab) - 1)
    End If
    
End Function

Public Function Online_XML1(ByVal asProc As String, ByVal asSpcno As String) As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    Online_XML1 = ""
    sFileName = "IPU2-" & asProc
    
    sParam = Select_Param1(asProc, asSpcno)
    
    sRetStr = Online_XML_Qry1(asProc, sParam)
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    Dim xDoc As MSXML.DOMDocument
    Set xDoc = New MSXML.DOMDocument
    If xDoc.Load(App.Path & "\XML\" & sFileName & ".xml") Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_S03
            Clear_XML_PInfo1
            display_online_parsing_PatInfo1 xDoc.childNodes, 0
            
        Case gXml_S07
            Clear_XML_Exam1
            display_online_parsing_ExamCode1 xDoc.childNodes, 0
            
        Case gXml_S08
            Clear_XML_Exam1
            display_online_parsing_ExamCode1 xDoc.childNodes, 0
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

        SaveXML_Data strErrText, sFileName
    End If
    
    
    Set xPE = Nothing
    Set xDoc = Nothing
    
    If sRetStr = "" Then
        Online_XML1 = 0
    Else
        Select Case asProc
        Case gXml_S03
            If gPat_Info_Select1.RERUN = "N" Then
                Online_XML1 = 1
            Else
                Online_XML1 = 0
            End If
        Case gXml_S07, gXml_S08
            If gExam_Select1(0).OK = 1 Then
                Online_XML1 = 1
            Else
                Online_XML1 = 0
            End If
        Case Else
            If InStr(1, gOnline_Ret1, vbTab) > 0 Then
                Online_XML1 = Left(gOnline_Ret1, InStr(1, gOnline_Ret1, vbTab) - 1)
            End If
        End Select
    End If
End Function

Public Function Online_XML_Qry1(ByVal asStrDiv As String, ByVal asParam As String) As String
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

    SaveXML_Data "[Use Proc => " & strDiv & " ]" & sParam, "IPU2_Info"
    send = oSOAP.wsLISInterface(strDiv, sParam)
    SaveXML_Data "[Return Proc => " & strDiv & " ]" & send, "IPU2_Info"
    Online_XML_Qry1 = send
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

Private Function Select_Param1(ByVal asProc As String, ByVal asSpcno As String) As String
    Dim sProc As String
    Dim sParam As String
    
    Select_Param1 = ""
    sProc = asProc
    
    Select Case sProc
    Case gXml_S01, gXml_S02, gXml_S03, gXml_S07, gXml_S08, gXml_S10
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
    
    Select_Param1 = sParam
    
End Function

Private Function TLA_Param1(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String
    Dim sProc As String
    Dim sParam As String
    
    TLA_Param1 = ""
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
    
    TLA_Param1 = sParam
    
End Function


'XML File Parsing===========================================================================================================
Public Sub display_online_parsing_ExamCode1(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
'            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "TST_CD"
                giIndex1 = giIndex1 + 1
                ReDim Preserve gExam_Select1(giIndex1)
                gExam_Select1(giIndex1).TST_CD = xNode.nodeValue
                gExam_Select1(giIndex1).TST_CNT = giIndex + 1
            End Select
            gExam_Select1(0).OK = 1
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_ExamCode1 xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub display_online_parsing_PatInfo1(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            Select Case xNode.parentNode.nodeName
            Case "TST_CD"
                gPat_Info_Select1.TST_CD = xNode.nodeValue
                gPat_Info_Select1.OK = 1
            Case "ACPT_DTETM"
                gPat_Info_Select1.ACPT_DTETM = xNode.nodeValue
            Case "ACPTNO_1"
                gPat_Info_Select1.ACPTNO_1 = xNode.nodeValue
            Case "AGE"
                gPat_Info_Select1.Age = xNode.nodeValue
            Case "HSP_CLS"
                gPat_Info_Select1.HSP_CLS = xNode.nodeValue
            Case "ORD_SITE"
                gPat_Info_Select1.ORD_SITE = xNode.nodeValue
            Case "PT_NM"
                gPat_Info_Select1.PT_NM = xNode.nodeValue
            Case "PT_NO"
                gPat_Info_Select1.PT_NO = xNode.nodeValue
            Case "RERUN"
                gPat_Info_Select1.RERUN = xNode.nodeValue
            Case "SEX"
                gPat_Info_Select1.Sex = xNode.nodeValue
            Case "SPC_CD_1"
                gPat_Info_Select1.SPC_CD_1 = xNode.nodeValue
            Case "TST_CLS"
                gPat_Info_Select1.TST_CLS = xNode.nodeValue
            Case "TST_FRCT_CD"
                gPat_Info_Select1.TST_FRCT_CD = xNode.nodeValue
            Case "TST_FRCT_CD1"
                gPat_Info_Select1.TST_FRCT_CD1 = xNode.nodeValue
            Case "TST_NM"
                gPat_Info_Select1.TST_NM = xNode.nodeValue
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_PatInfo1 xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub display_online_parsing_TLAInfo1(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "TST_DTE"
                giIndex1 = giIndex1 + 1
                ReDim Preserve gTLA_Info_Select1(giIndex1)
                
                gTLA_Info_Select1(giIndex1).TST_DTE = xNode.nodeValue
            Case "SPCNO"
                gTLA_Info_Select1(giIndex1).SPCNO = xNode.nodeValue
            End Select
            gTLA_Info_Select1(giIndex1).OK = 1
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_TLAInfo1 xNode.childNodes, Indent
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
'End Function
'
'Public Function Online_Result_Qry(ByVal asParam As String) As String
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
'    sParam = asParam
'
'    SaveXML_Data "[Save Result]" & sParam
'    send = oSOAP.wsLISInterface(strDiv, sParam)
'    SaveXML_Data "[Save Result => Return]" & send
'    Online_Result_Qry = send
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

Public Function Online_Result_Qry1(ByVal asParam As String) As String
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
    
    SaveXML_Data "[Save Result]" & sParam, "IPU2_Res"
    send = oSOAP.wsLISInterface(strDiv, sParam)
    SaveXML_Data "[Save Result => Return]" & send, "IPU2_Res"
    Online_Result_Qry1 = send
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
Public Sub SaveXML_Data1(argSQL As String)
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

Public Sub Xml_Log1(argSQL As String, argFileName As String)
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

