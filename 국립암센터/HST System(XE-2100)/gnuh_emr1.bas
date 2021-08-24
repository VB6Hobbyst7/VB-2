Attribute VB_Name = "gnuh_emr1"
'Option Explicit
'
'Type Order_Select
'    SPC_NO      As String
'    PT_NO       As String
'    PT_NM       As String
'    ACPT_DTE    As String
'    ACPT_NO     As String
'    TST_CD      As String
'    WRK_UNT     As String
'    TST_DTE     As String
'    TST_STAT    As String
'    WD_NO       As String
'    SPC_NM      As String
'    SPC_CD      As String
'    ok          As Integer
'End Type
'Public gOrder_Select1 As Order_Select
'Public gOrder_List1() As Order_Select
'Public gWork_Select1() As Order_Select
'Public giIndex1  As Long
'
'Type Patient_Info
'    PTNO        As String
'    PATNAME     As String
'    Sex         As String
'    Age         As String
'    DPCD        As String
'    WD_NO       As String
'    SPC_CD      As String
'    SPC_NM      As String
'    ACPT_NO     As String
'    ACPT_DTM    As String
'    TST_STAT    As String
'    ok          As Integer
'End Type
'Public gPatient_Info1 As Patient_Info
'
'Public gOnline_Ret1 As String
'
'Public Sub SaveXMLFile1(argXML As String)
''argSQL의 내용을 파일로 저장
'    Dim FilNum
'
'    FilNum = FreeFile
'
'    If Dir(App.Path & "\ipu2.xml") <> "" Then
'        Kill App.Path & "\ipu2.xml"
'    End If
'
'    Open App.Path & "\ipu2.xml" For Append As FilNum
'    Print #FilNum, argXML
'    Close FilNum
'
'End Sub
'
''2009.10.01 윤영기
''검체번호로 인터페이스 하지 않은 검사코드 가져오기
''return : 1 => 검사 존재, 0 => 검사 없음
''gOrder_select에 파라미터 저장
'Public Function Get_Order_1(asSID) As Integer
'    Dim sRetStr As String
'
'    gOrder_Select1.ok = 0
'
'    giIndex1 = -1
'    ReDim gOrder_List1(0)
'
'    sRetStr = Get_Qry_OrderList1(asSID)
'
'    'SaveXMLFile1 sRetStr
'    Save_Xml_Data1 sRetStr, "ipu2_orderlist"
'
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    'If xDoc.Load(App.Path & "\ipu2.xml") Then
'    If xDoc.Load(App.Path & "\Res" & "\ipu2_orderlist.xml") Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_Order_Parsing1 xDoc.childNodes, 0
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
'    Get_Order_1 = gOrder_Select1.ok
'End Function
'
''PG_SRL.INTERFACE_S21
''인터페이스 웹서버에서 연속 검사 데이타 가져오기
'Public Function Get_Qry_OrderList_M(ByVal asSID As String) As String
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
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
'
'    strDiv = "PG_SRL.INTERFACE_S21"
'    'asSID = "09092251028"
'
'    sParam = "<Table>" & _
'                      "<QID><![CDATA[PG_SRL.INTERFACE_S21]]></QID>" & _
'                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'                      "<USERID><![CDATA[LIA]]></USERID>" & _
'                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'                      "<P0><![CDATA[" & asSID & "]]></P0>" & _
'                      "<P1><![CDATA[" & "" & "]]></P1>" & _
'               "</Table>"
'
'
'    send = oSOAP.wsLISInterface(strDiv, sParam)
'
'    Get_Qry_OrderList_M = send
'
'    Set oSOAP = Nothing
'
'    DoEvents
'
'    Exit Function
'
'ErrHandle:
'    If oSOAP.FaultString <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
'    End If
'    If Trim(Err.Description) <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
'    End If
'
'    Set oSOAP = Nothing
'
'End Function
'
'
''PG_SRL.INTERFACE_S03
''인터페이스 웹서버에서 데이타 가져오기
'Public Function Get_Qry_OrderList1(ByVal asSID As String) As String
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
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
'
'    strDiv = "PG_SRL.INTERFACE_S03"
'    'asSID = "09092251028"
'
'    sParam = "<Table>" & _
'                      "<QID><![CDATA[PG_SRL.INTERFACE_S03]]></QID>" & _
'                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'                      "<USERID><![CDATA[LIA]]></USERID>" & _
'                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'                      "<P0><![CDATA[" & asSID & "]]></P0>" & _
'                      "<P1><![CDATA[" & "" & "]]></P1>" & _
'               "</Table>"
'
'
'    send = oSOAP.wsLISInterface("PG_SRL.INTERFACE_S03", sParam)
'
'    Get_Qry_OrderList1 = send
'
'    Set oSOAP = Nothing
'
'    DoEvents
'
'    Exit Function
'
'ErrHandle:
'    If oSOAP.FaultString <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
'    End If
'    If Trim(Err.Description) <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
'    End If
'
'    Set oSOAP = Nothing
'
'End Function
'
'Public Sub Display_Order_Parsing_M(ByRef Nodes As MSXML.IXMLDOMNodeList, _
'    ByVal Indent As Integer)
'
'    Dim xNode As MSXML.IXMLDOMNode
'    Indent = Indent + 2
'
'    For Each xNode In Nodes
'        If xNode.nodeType = 4 Then
'        'If xNode.nodeType = NODE_TEXT Then
'        'If xNode.nodeType = NODE_ATTRIBUTE Then
'        'If xNode.nodeType = NODE_ELEMENT Then
'            Select Case xNode.parentNode.nodeName
'            Case "SPC_NO"
'                giIndex = giIndex + 1
'                ReDim Preserve gOrder_List(giIndex)
'                gOrder_Select.ok = giIndex + 1
'            Case "TST_CD"
'                gOrder_List(giIndex).TST_CD = xNode.nodeValue
'            Case "CSUBCD3_NM"
'                gOrder_List(giIndex).SPC_LAST = xNode.nodeValue
'
'            End Select
'
'        End If
'
'        If xNode.hasChildNodes Then
'            Display_Order_Parsing_M xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub
'
'
''XML File Parsing
'Public Sub Display_Order_Parsing1(ByRef Nodes As MSXML.IXMLDOMNodeList, _
'    ByVal Indent As Integer)
'
'    Dim xNode As MSXML.IXMLDOMNode
'    Indent = Indent + 2
'
'    For Each xNode In Nodes
'        If xNode.nodeType = 4 Then
'        'If xNode.nodeType = NODE_TEXT Then
'        'If xNode.nodeType = NODE_ATTRIBUTE Then
'        'If xNode.nodeType = NODE_ELEMENT Then
'            Select Case xNode.parentNode.nodeName
'            Case "PT_NO"
'                giIndex1 = giIndex1 + 1
'                ReDim Preserve gOrder_List1(giIndex1)
'
'                gOrder_List1(giIndex1).PT_NO = xNode.nodeValue
'            Case "ACPT_DTE": gOrder_List1(giIndex1).ACPT_DTE = xNode.nodeValue
'            Case "ACPT_NO":  gOrder_List1(giIndex1).ACPT_NO = xNode.nodeValue
'            Case "TST_CD":   gOrder_List1(giIndex1).TST_CD = xNode.nodeValue
'            Case "WRK_UNT":  gOrder_List1(giIndex1).WRK_UNT = xNode.nodeValue
'            Case "PT_NM":    gOrder_List1(giIndex1).PT_NM = xNode.nodeValue
'            End Select
'            gOrder_Select1.ok = 1
'        End If
'
'        If xNode.hasChildNodes Then
'            Display_Order_Parsing1 xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub
'
'
''2009.10.01 윤영기
''날짜, 검사코드로 검사리스트 가져오기
''return : 1 => 검사 존재, 0 => 검사 없음
''gOrder_select에 파라미터 저장
'Public Function Get_WorkList1(asFromDT As String, asToDT As String, asTest As String, asGubun As String) As Integer
'    Dim sRetStr As String
'
'    ReDim Preserve gWork_Select1(0)
'    giIndex1 = -1
'
'    sRetStr = Get_Qry_WorkList1(asFromDT, asToDT, asTest, asGubun)
'
'    SaveXMLFile1 sRetStr
'
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    If xDoc.Load(App.Path & "\ipu2.xml") Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_Work_Parsing1 xDoc.childNodes, 0
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
'    Get_WorkList1 = gOrder_Select1.ok
'End Function
'
'Public Function Get_WorkList11(asFromDT As String, asToDT As String, asTest As String, asGubun As String) As Integer
'    Dim sRetStr As String
'
'    ReDim Preserve gWork_Select1(0)
'    giIndex1 = -1
'
'    sRetStr = Get_Qry_WorkList11(asFromDT, asToDT, asTest, asGubun)
'
'    SaveXMLFile1 sRetStr
'
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    If xDoc.Load(App.Path & "\ipu2.xml") Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_Work_Parsing11 xDoc.childNodes, 0
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
'    Get_WorkList11 = gOrder_Select1.ok
'End Function
'
''PG_SRL.INTERFACE_S03
''인터페이스 웹서버에서 데이타 가져오기
'Public Function Get_Qry_WorkList1(asFromDT As String, asToDT As String, asTest As String, asGubun As String) As String
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
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
'
'    strDiv = "PG_SRL.INTERFACE_S13"
'    'asSID = "09092251028"
'
'    sParam = "<Table>" & _
'                      "<QID><![CDATA[PG_SRL.INTERFACE_S13]]></QID>" & _
'                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'                      "<USERID><![CDATA[LIA]]></USERID>" & _
'                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'                      "<P0><![CDATA[" & asFromDT & "]]></P0>" & _
'                      "<P1><![CDATA[" & asToDT & "]]></P1>" & _
'                      "<P2><![CDATA[" & asTest & "]]></P2>" & _
'                      "<P3><![CDATA[" & asGubun & "]]></P3>" & _
'                      "<P4><![CDATA[" & "" & "]]></P4>" & _
'               "</Table>"
'
'    'Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
'
'    send = oSOAP.wsLISInterface(strDiv, sParam)
'
'    Get_Qry_WorkList1 = send
'
'    Set oSOAP = Nothing
'
'    DoEvents
'
'    Exit Function
'
'ErrHandle:
'    If oSOAP.FaultString <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
'    End If
'    If Trim(Err.Description) <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
'    End If
'
'    Set oSOAP = Nothing
'
'End Function
'
'Public Function Get_Qry_WorkList11(asFromDT As String, asToDT As String, asTest As String, asGubun As String) As String
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
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
'
'    strDiv = "PG_SRL.INTERFACE_S15"
'    'asSID = "09092251028"
'
'    sParam = "<Table>" & _
'                      "<QID><![CDATA[PG_SRL.INTERFACE_S15]]></QID>" & _
'                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'                      "<USERID><![CDATA[LIA]]></USERID>" & _
'                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'                      "<P0><![CDATA[" & asFromDT & "]]></P0>" & _
'                      "<P1><![CDATA[" & asToDT & "]]></P1>" & _
'                      "<P2><![CDATA[" & asTest & "]]></P2>" & _
'                      "<P3><![CDATA[" & asGubun & "]]></P3>" & _
'                      "<P4><![CDATA[" & "" & "]]></P4>" & _
'               "</Table>"
'
'    'Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
'
'    send = oSOAP.wsLISInterface(strDiv, sParam)
'
'    Get_Qry_WorkList11 = send
'
'    Set oSOAP = Nothing
'
'    DoEvents
'
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
''XML File Parsing
'Public Sub Display_Work_Parsing1(ByRef Nodes As MSXML.IXMLDOMNodeList, _
'    ByVal Indent As Integer)
'
'    Dim xNode As MSXML.IXMLDOMNode
'    Indent = Indent + 2
'
'    For Each xNode In Nodes
'        'Debug.Print xNode.nodeType
'        'Debug.Print xNode.nodeType & vbTab & xNode.parentNode.nodeName & " : " & xNode.nodeValue
'        If xNode.nodeType = 4 Then
'        'If xNode.nodeType = NODE_TEXT Then
'        'If xNode.nodeType = NODE_ATTRIBUTE Then
'        'If xNode.nodeType = NODE_ELEMENT Then
'            'Debug.Print xNode.parentNode.nodeName & " : " & xNode.nodeValue
'
'            Select Case xNode.parentNode.nodeName
'            Case "PT_NO":    gWork_Select1(giIndex1).PT_NO = xNode.nodeValue
'            Case "PT_NM":    gWork_Select1(giIndex1).PT_NM = xNode.nodeValue
'            Case "SPC_NO"
'                giIndex1 = giIndex1 + 1
'                ReDim Preserve gWork_Select1(giIndex1)
'                gWork_Select1(giIndex1).SPC_NO = xNode.nodeValue
'            Case "TST_DTE":  gWork_Select1(giIndex1).TST_DTE = xNode.nodeValue
'            Case "ACPT_NO":  gWork_Select1(giIndex1).ACPT_NO = xNode.nodeValue
'            Case "TST_CD":   gWork_Select1(giIndex1).TST_CD = xNode.nodeValue
'            Case "TST_STAT": gWork_Select1(giIndex1).TST_STAT = xNode.nodeValue
'            Case "WD_NO":    gWork_Select1(giIndex1).WD_NO = xNode.nodeValue
'            Case "SPC_NM":   gWork_Select1(giIndex1).SPC_NM = xNode.nodeValue
'
'            End Select
'            gOrder_Select1.ok = 1
'        End If
'
'        If xNode.hasChildNodes Then
'
'
'            Display_Work_Parsing1 xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub
'
''XML File Parsing
'Public Sub Display_Work_Parsing11(ByRef Nodes As MSXML.IXMLDOMNodeList, _
'    ByVal Indent As Integer)
'
'    Dim xNode As MSXML.IXMLDOMNode
'    Indent = Indent + 2
'
'    For Each xNode In Nodes
'        'Debug.Print xNode.nodeType
'        Debug.Print xNode.nodeType & vbTab & xNode.parentNode.nodeName & " : " & xNode.nodeValue
'        If xNode.nodeType = 4 Then
'        'If xNode.nodeType = NODE_TEXT Then
'        'If xNode.nodeType = NODE_ATTRIBUTE Then
'        'If xNode.nodeType = NODE_ELEMENT Then
'            'Debug.Print xNode.parentNode.nodeName & " : " & xNode.nodeValue
'
'            Select Case xNode.parentNode.nodeName
'            Case "PT_NO":    gWork_Select1(giIndex1).PT_NO = xNode.nodeValue
'            Case "PT_NM":    gWork_Select1(giIndex1).PT_NM = xNode.nodeValue
'            Case "SPC_NO":   gWork_Select1(giIndex1).SPC_NO = xNode.nodeValue
'            Case "TST_DTE":  gWork_Select1(giIndex1).TST_DTE = xNode.nodeValue
'            Case "ACPT_NO":  gWork_Select1(giIndex1).ACPT_NO = xNode.nodeValue
'            Case "TST_CD":   gWork_Select1(giIndex1).TST_CD = xNode.nodeValue
'            Case "TST_STAT": gWork_Select1(giIndex1).TST_STAT = xNode.nodeValue
'            Case "WD_NO":    gWork_Select1(giIndex1).WD_NO = xNode.nodeValue
'            Case "SPC_NO":   gWork_Select1(giIndex1).SPC_NO = xNode.nodeValue
'            Case "SPC_NM":   gWork_Select1(giIndex1).SPC_NM = xNode.nodeValue
'
'            End Select
'            gOrder_Select1.ok = 1
'        End If
'
'        If xNode.hasChildNodes Then
'            giIndex1 = giIndex1 + 1
'            ReDim Preserve gWork_Select1(giIndex1)
'
'            Display_Work_Parsing11 xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub
'
'Public Sub Clear_PatInfo1()
'    gPatient_Info1.PTNO = ""
'    gPatient_Info1.PATNAME = ""
'    gPatient_Info1.Sex = ""
'    gPatient_Info1.Age = ""
'    gPatient_Info1.WD_NO = ""
'    gPatient_Info1.SPC_CD = ""
'    gPatient_Info1.SPC_NM = ""
'    gPatient_Info1.ACPT_NO = ""
'    gPatient_Info1.ACPT_DTM = ""
'    gPatient_Info1.TST_STAT = ""
'End Sub
'
''2009.10.01 윤영기
''검체번호로 환자정보 가져오기
''return : 1 => 검사 존재, 0 => 검사 없음
''gOrder_select에 파라미터 저장
'Public Function Get_PatInfo1(asSID) As Integer
'    Dim sRetStr As String
'
'    gPatient_Info1.ok = 0
'
'    Clear_PatInfo1
'
'    sRetStr = Get_Qry_PatInfo1(asSID)
'
'    'SaveXMLFile1 sRetStr
'    Save_Xml_Data1 sRetStr, "ipu2_patinfo"
'
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    'If xDoc.Load(App.Path & "\ipu2.xml") Then
'    If xDoc.Load(App.Path & "\Res" & "\ipu2_patinfo.xml") Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_PatInfo_Parsing1 xDoc.childNodes, 0
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
'    Get_PatInfo1 = gPatient_Info1.ok
'End Function
'
''PG_SRL.INTERFACE_S06
''인터페이스 웹서버에서 데이타 가져오기
'Public Function Get_Qry_PatInfo1(ByVal asSID As String) As String
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
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
'
'    strDiv = "PG_SRL.INTERFACE_S06"
'    'asSID = "09092251028"
'
'    sParam = "<Table>" & _
'                      "<QID><![CDATA[PG_SRL.INTERFACE_S06]]></QID>" & _
'                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'                      "<USERID><![CDATA[LIA]]></USERID>" & _
'                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'                      "<P0><![CDATA[" & asSID & "]]></P0>" & _
'                      "<P1><![CDATA[" & "" & "]]></P1>" & _
'               "</Table>"
'
''    Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
'
'    send = oSOAP.wsLISInterface(strDiv, sParam)
'
'    Get_Qry_PatInfo1 = send
'
'    Set oSOAP = Nothing
'
'    DoEvents
'
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
''XML File Parsing
'Public Sub Display_PatInfo_Parsing1(ByRef Nodes As MSXML.IXMLDOMNodeList, _
'    ByVal Indent As Integer)
'
'    Dim xNode As MSXML.IXMLDOMNode
'    Indent = Indent + 2
'
'    For Each xNode In Nodes
'        If xNode.nodeType = 4 Then
'        'If xNode.nodeType = NODE_TEXT Then
'        'If xNode.nodeType = NODE_ATTRIBUTE Then
'        'If xNode.nodeType = NODE_ELEMENT Then
'            Select Case xNode.parentNode.nodeName
'            Case "PTNO":     gPatient_Info1.PTNO = xNode.nodeValue
'            Case "PATNAME":  gPatient_Info1.PATNAME = xNode.nodeValue
'            Case "SEX":      gPatient_Info1.Sex = xNode.nodeValue
'            Case "AGE":      gPatient_Info1.Age = xNode.nodeValue
'            Case "WD_NO":    gPatient_Info1.WD_NO = xNode.nodeValue
'            Case "SPC_CD":   gPatient_Info1.SPC_CD = xNode.nodeValue
'            Case "SPC_NM":   gPatient_Info1.SPC_NM = xNode.nodeValue
'            Case "ACPT_NO":  gPatient_Info1.ACPT_NO = xNode.nodeValue
'            Case "ACPT_DTM": gPatient_Info1.ACPT_DTM = xNode.nodeValue
'            Case "TST_STAT": gPatient_Info1.TST_STAT = xNode.nodeValue
'            End Select
'            gPatient_Info1.ok = 1
'        End If
'
'        If xNode.hasChildNodes Then
'            Display_PatInfo_Parsing1 xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub
'
'Public Function Online_Result1(ByVal asSpcno As String, _
'                              ByVal asExam As String, _
'                              ByVal asRes As String, _
'                              ByVal asEquip As String, _
'                              ByVal asCount As String, _
'                              ByVal asEqFlag As String) As String
'
'
'    Dim sRetStr As String
'
'
'    Online_Result1 = ""
'
'    gOnline_Ret1 = ""
'
'    sRetStr = Online_Result_Qry(asSpcno, asExam, asRes, asEquip, asCount, asEqFlag)
'
'    'SaveXMLFile1 sRetStr
'    Save_Xml_Data sRetStr, "ipu2_result"
'
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    'If xDoc.Load(App.Path & "\ipu2.xml") Then
'    If xDoc.Load(App.Path & "\Res" & "\ipu2_result.xml") Then
'    'If xDoc.Load(sRetStr) Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_Online_Parsing1 xDoc.childNodes, 0
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
'    If InStr(1, gOnline_Ret1, vbTab) > 0 Then
'        Online_Result1 = Left(gOnline_Ret1, InStr(1, gOnline_Ret1, vbTab) - 1)
'    End If
'
'End Function
'
'
'
'
'Public Function Online_Result_Qry1(ByVal asSpcno As String, _
'                              ByVal asExam As String, _
'                              ByVal asRes As String, _
'                              ByVal asEquip As String, _
'                              ByVal asCount As String, _
'                              ByVal asEqFlag As String) As String
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
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
'
'    strDiv = "PG_SRL.INTERFACE_I01"
'    'asSID = "09092251028"
'
'    sParam = "<Table>" & _
'                      "<QID><![CDATA[PG_SRL.INTERFACE_I01]]></QID>" & _
'                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'                      "<USERID><![CDATA[LIA]]></USERID>" & _
'                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'                      "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
'                      "<P1><![CDATA[" & asExam & "]]></P1>" & _
'                      "<P2><![CDATA[" & asRes & "]]></P2>" & _
'                      "<P3><![CDATA[" & asEqFlag & "]]></P3>" & _
'                      "<P4><![CDATA[" & asEquip & "]]></P4>" & _
'                      "<P5><![CDATA[]]></P5>" & _
'                      "<P6><![CDATA[" & asCount & "]]></P6>" & _
'                      "<P7><![CDATA[]]></P7>" & _
'                      "<P8><![CDATA[]]></P8>" & _
'                      "<P9><![CDATA[]]></P9>" & _
'                      "<P10><![CDATA[]]></P10>" & _
'               "</Table>"
'
''    Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
'
'    SaveData "[Save Result]" & sParam
'
'    send = oSOAP.wsLISInterface(strDiv, sParam)
'
'    SaveData "[Save Result => Return]" & send
'
'    Online_Result_Qry1 = send
'
'    Set oSOAP = Nothing
'
'    DoEvents
'
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
''XML File Parsing
'Public Sub Display_Online_Parsing1(ByRef Nodes As MSXML.IXMLDOMNodeList, _
'    ByVal Indent As Integer)
'
'    Dim xNode As MSXML.IXMLDOMNode
'    Indent = Indent + 2
'
'    For Each xNode In Nodes
'
'        If xNode.nodeType = 4 Then
'            gOnline_Ret1 = gOnline_Ret1 & xNode.nodeValue & vbTab
'        End If
'
'        If xNode.hasChildNodes Then
'            Display_Online_Parsing1 xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub
'
''2009.10.01 윤영기
''QC검체번호로 인터페이스 하지 않은 검사코드 가져오기
''return : 1 => 검사 존재, 0 => 검사 없음
''gOrder_select에 파라미터 저장
'Public Function Get_QCOrder1(asSID, asGubun As String) As Integer
'    Dim sRetStr As String
'
'    gOrder_Select1.ok = 0
'
'    giIndex1 = -1
'    ReDim gOrder_List1(0)
'
'    sRetStr = Get_Qry_QCOrderList1(asSID, asGubun)
'
'    'SaveXMLFile1 sRetStr
'    Save_Xml_Data sRetStr, "ipu2_qcorder"
'
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    'If xDoc.Load(App.Path & "\ipu2.xml") Then
'    If xDoc.Load(App.Path & "\Res" & "\ipu2_qcorder.xml") Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_QCOrder_Parsing1 xDoc.childNodes, 0
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
'    Get_QCOrder1 = gOrder_Select1.ok
'End Function
'
''PG_SRL.INTERFACE_S17
''인터페이스 웹서버에서 데이타 가져오기
'Public Function Get_Qry_QCOrderList1(ByVal asSID As String, ByVal asGubun As String) As String
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
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
'
'    strDiv = "PG_SRL.INTERFACE_S17"
'    'asSID = "09092251028"
'
'    sParam = "<Table>" & _
'                      "<QID><![CDATA[PG_SRL.INTERFACE_S17]]></QID>" & _
'                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'                      "<USERID><![CDATA[LIA]]></USERID>" & _
'                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'                      "<P0><![CDATA[" & asGubun & "]]></P0>" & _
'                      "<P1><![CDATA[" & asSID & "]]></P1>" & _
'                      "<P2><![CDATA[" & "" & "]]></P2>" & _
'               "</Table>"
'
'
'    send = oSOAP.wsLISInterface("PG_SRL.INTERFACE_S03", sParam)
'
'    Get_Qry_QCOrderList1 = send
'
'    Set oSOAP = Nothing
'
'    DoEvents
'
'    Exit Function
'
'ErrHandle:
'    If oSOAP.FaultString <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
'    End If
'    If Trim(Err.Description) <> "" Then
'        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
'    End If
'
'    Set oSOAP = Nothing
'
'End Function
'
''XML File Parsing
'Public Sub Display_QCOrder_Parsing1(ByRef Nodes As MSXML.IXMLDOMNodeList, _
'    ByVal Indent As Integer)
'
'    Dim xNode As MSXML.IXMLDOMNode
'    Indent = Indent + 2
'
'    For Each xNode In Nodes
'        If xNode.nodeType = 4 Then
'        'If xNode.nodeType = NODE_TEXT Then
'        'If xNode.nodeType = NODE_ATTRIBUTE Then
'        'If xNode.nodeType = NODE_ELEMENT Then
'            Select Case xNode.parentNode.nodeName
'            Case "INST_DTM"
'                giIndex1 = giIndex1 + 1
'                ReDim Preserve gOrder_List(giIndex1)
'
'                gOrder_List1(giIndex1).ACPT_DTE = xNode.nodeValue
'
'            Case "TST_CD":    gOrder_List1(giIndex1).TST_CD = xNode.nodeValue
'            Case "TST_NM"
'            Case "EQUIP_CD":  gOrder_List1(giIndex1).ACPT_NO = xNode.nodeValue
'            Case "CTRL_CD":   gOrder_List1(giIndex1).PT_NO = xNode.nodeValue
'            Case "LOT_NO":    gOrder_List1(giIndex1).PT_NM = xNode.nodeValue
'            Case "BARCODE_CD": gOrder_List1(giIndex1).SPC_NO = xNode.nodeValue
'            End Select
'            gOrder_Select1.ok = 1
'        End If
'
'        If xNode.hasChildNodes Then
'            Display_QCOrder_Parsing1 xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub
'
'
'Public Function QCOnline_Result1(ByVal asSpcno As String, _
'                                ByVal asLotno As String, _
'                              ByVal asExam As String, _
'                              ByVal asRes As String, _
'                              ByVal asEquip As String, _
'                              ByVal asCount As String, _
'                              ByVal asResDate As String) As String
'
'
'    Dim sRetStr As String
'
'
'    QCOnline_Result1 = ""
'
'    gOnline_Ret1 = ""
'
'    sRetStr = QCOnline_Result_Qry1(asSpcno, asLotno, asExam, asRes, asEquip, asCount, asResDate)
'
'    SaveXMLFile1 sRetStr
'
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    If xDoc.Load(App.Path & "\ipu2.xml") Then
'    'If xDoc.Load(sRetStr) Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_Online_Parsing1 xDoc.childNodes, 0
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
'    If InStr(1, gOnline_Ret1, vbTab) > 0 Then
'        QCOnline_Result1 = Left(gOnline_Ret1, InStr(1, gOnline_Ret1, vbTab) - 1)
'    End If
'
'End Function
'
'
'
'
'Public Function QCOnline_Result_Qry1(ByVal asSpcno As String, _
'                              ByVal asLotno As String, _
'                              ByVal asExam As String, _
'                              ByVal asRes As String, _
'                              ByVal asEquip As String, _
'                              ByVal asCount As String, _
'                              ByVal asResDate As String) As String
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
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
'
'    strDiv = "PG_SRL.INTERFACE_U04"
'    'asSID = "09092251028"
'
'    sParam = "<Table>" & _
'                      "<QID><![CDATA[PG_SRL.INTERFACE_U04]]></QID>" & _
'                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'                      "<USERID><![CDATA[LIA]]></USERID>" & _
'                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'                      "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
'                      "<P1><![CDATA[" & asEquip & "]]></P1>" & _
'                      "<P1><![CDATA[" & asLotno & "]]></P1>" & _
'                      "<P2><![CDATA[" & asResDate & "]]></P2>" & _
'                      "<P3><![CDATA[" & asCount & "]]></P3>" & _
'                      "<P4><![CDATA[" & asEquip & "]]></P4>" & _
'                      "<P5><![CDATA[" & asExam & "]]></P5>" & _
'                      "<P6><![CDATA[" & asRes & "]]></P6>" & _
'                      "<P7><![CDATA[]]></P7>" & _
'                      "<P8><![CDATA[]]></P8>" & _
'               "</Table>"
'
''    Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
'
'    SaveData "[Save Result]" & sParam
'
'    send = oSOAP.wsLISInterface(strDiv, sParam)
'
'    SaveData "[Save Result => Return]" & send
'
'    QCOnline_Result_Qry1 = send
'
'    Set oSOAP = Nothing
'
'    DoEvents
'
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
'Public Sub Save_Xml_Data1(argSQL As String, argFileName As String)
''argSQL의 내용을 파일로 저장
'    Dim FilNum
'    Dim sFileName As String
'
'    FilNum = FreeFile
'
'    If Dir(App.Path & "\" & "Res", vbDirectory) <> "Res" Then
'        MkDir (App.Path & "\" & "Res")
'    End If
'
''    sFileName = Format(CDate(frmMain.txtToday.Text), "yyyymmdd")
'    sFileName = argFileName
'    If Dir(App.Path & "\" & "Res" & "\" & sFileName & ".xml") <> "" Then
'        Kill App.Path & "\" & "Res" & "\" & sFileName & ".xml"
'    End If
'
'    Open App.Path & "\" & "Res" & "\" & sFileName & ".xml" For Append As FilNum
'    Print #FilNum, argSQL
'    Close FilNum
'End Sub
'
''New Result Trans sub start I08 사용 ==========================================================================
'Public Function Online_Result_New1(ByVal asSpcno As String, _
'                              ByVal asExam As String, _
'                              ByVal asRes As String, _
'                              ByVal asEquip As String, _
'                              ByVal asCount As String, _
'                              ByVal asEqFlag As String, _
'                              ByVal asUser As String) As String
'
'
'    Dim sRetStr As String
'
'
'    Online_Result_New1 = ""
'
'    gOnline_Ret1 = ""
'
'    sRetStr = Online_Result_Qry_New1(asSpcno, asExam, asRes, asEquip, asCount, asEqFlag, asUser)
'
'    'SaveXMLFile sRetStr
'    Save_Xml_Data sRetStr, "ipu2_result"
'
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    If xDoc.Load(App.Path & "\Res\ipu2_result.xml") Then
'    'If xDoc.Load(sRetStr) Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_Online_Parsing1 xDoc.childNodes, 0
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
'    If InStr(1, gOnline_Ret1, vbTab) > 0 Then
'        Online_Result_New1 = Left(gOnline_Ret1, InStr(1, gOnline_Ret1, vbTab) - 1)
'    End If
'
'End Function
'
'Public Function Online_Result_Qry_New1(ByVal asSpcno As String, _
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
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
'
'    strDiv = "PG_SRL.INTERFACE_I08"
'    'asSID = "09092251028"
'
'    sParam = "<Table>" & _
'                      "<QID><![CDATA[PG_SRL.INTERFACE_I08]]></QID>" & _
'                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
'                      "<USERID><![CDATA[LIA]]></USERID>" & _
'                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
'                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
'                      "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
'                      "<P1><![CDATA[" & asExam & "]]></P1>" & _
'                      "<P2><![CDATA[" & asRes & "]]></P2>" & _
'                      "<P3><![CDATA[" & asEqFlag & "]]></P3>" & _
'                      "<P4><![CDATA[" & asEquip & "]]></P4>" & _
'                      "<P5><![CDATA[]]></P5>" & _
'                      "<P6><![CDATA[" & asCount & "]]></P6>" & _
'                      "<P7><![CDATA[]]></P7>" & _
'                      "<P8><![CDATA[]]></P8>" & _
'                      "<P9><![CDATA[" & asUser & "]]></P9>" & _
'                      "<P10><![CDATA[]]></P10>" & _
'                      "<P11><![CDATA[]]></P11>" & _
'               "</Table>"
'
''    Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
'
'    SaveData "[Save Result]" & sParam
'
'    send = oSOAP.wsLISInterface(strDiv, sParam)
'
'    SaveData "[Save Result => Return]" & send
'
'    Online_Result_Qry_New1 = send
'
'    Set oSOAP = Nothing
'
'    DoEvents
'
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
''New Result Trans sub end I08 사용 ===========================================================================
'
'
'
'
