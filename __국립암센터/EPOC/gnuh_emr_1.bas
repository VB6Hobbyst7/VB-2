Attribute VB_Name = "gnuh_emr"
Option Explicit

Type Order_Select
    SPC_NO      As String
    PT_NO       As String
    PT_NM       As String
    ACPT_DTE    As String
    ACPT_NO     As String
    TST_CD      As String
    WRK_UNT     As String
    Sex         As String
    Age         As String
    TST_DTE     As String
    TST_STAT    As String
    WD_NO       As String
    SPC_NM      As String
    SPC_CD      As String
    ok          As Integer
End Type
Public gOrder_Select As Order_Select
Public gOrder_List() As Order_Select
Public gWork_Select() As Order_Select
Public giIndex  As Long

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
    ok          As Integer
End Type
Public gPatient_Info As Patient_Info

Type QC_Info
    INST_DTM    As String
    LOT_NO      As String
    TST_CD      As String
    equip_cd    As String
    CTRL_CD     As String
    LOT_NO1     As String
    BARCODE_CD  As String
    USE_YN      As String
    ok          As Integer
End Type
Public gQC_Info() As QC_Info

Public gResultExamCode() As String


Public gOnline_Ret As String
Public gOrderExam As String
Public gReceCode As String

Public Sub SaveXMLFile(argXML As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
        
    FilNum = FreeFile
    
    If Dir(App.Path & "\order.xml") <> "" Then
        Kill App.Path & "\order.xml"
    End If
    
    Open App.Path & "\order.xml" For Append As FilNum
    Print #FilNum, argXML
    Close FilNum
    
End Sub

'2009.10.01 윤영기
'검체번호로 인터페이스 하지 않은 검사코드 가져오기
'return : 1 => 검사 존재, 0 => 검사 없음
'gOrder_select에 파라미터 저장
Public Function Get_Order(asSID) As Integer
    Dim sRetStr As String
    
    gOrder_Select.ok = 0
    
    giIndex = -1
    ReDim gOrder_List(0)
    
    sRetStr = Get_Qry_OrderList(asSID)
    
    SaveXMLFile sRetStr
    
    Dim xDoc As MSXML.DOMDocument
    
    Set xDoc = New MSXML.DOMDocument
    
    gReceCode = ""
    If xDoc.Load(App.Path & "\order.xml") Then
        ' 문서가 성공적으로 로드되었습니다.
        ' 이제 재미있는 작업을 수행합니다.
        
        Display_Order_Parsing xDoc.childNodes, 0
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
        
        Save_Raw_Data strErrText
    End If
    
    Set xPE = Nothing

    Set xDoc = Nothing
    
    Get_Order = gOrder_Select.ok
End Function

'PG_SRL.INTERFACE_S03
'인터페이스 웹서버에서 데이타 가져오기
Public Function Get_Qry_OrderList(ByVal asSID As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    
    strDiv = "PG_SRL.INTERFACE_S07"
    'asSID = "09092251028"
    
    sParam = "<Table>" & _
                      "<QID><![CDATA[PG_SRL.INTERFACE_S07]]></QID>" & _
                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                      "<USERID><![CDATA[LIA]]></USERID>" & _
                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                      "<P0><![CDATA[" & asSID & "]]></P0>" & _
                      "<P1><![CDATA[" & "" & "]]></P1>" & _
               "</Table>"
                   
    
    send = oSOAP.wsLISInterface("PG_SRL.INTERFACE_S03", sParam)
    
    Get_Qry_OrderList = send
    
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
    
    Set oSOAP = Nothing
    
End Function

'XML File Parsing
Public Sub Display_Order_Parsing(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
        If xNode.nodeType = 4 Then
        'If xNode.nodeType = NODE_TEXT Then
        'If xNode.nodeType = NODE_ATTRIBUTE Then
        'If xNode.nodeType = NODE_ELEMENT Then
            Select Case xNode.parentNode.nodeName
'            Case "PT_NO"
'                giIndex = giIndex + 1
'                ReDim Preserve gOrder_List(giIndex)
'                ReDim Preserve gResultExamCode(giIndex)
                
'                gOrder_List(giIndex).PT_NO = xNode.nodeValue
'            Case "ACPT_DTE": gOrder_List(giIndex).ACPT_DTE = xNode.nodeValue
'            Case "ACPT_NO":  gOrder_List(giIndex).ACPT_NO = xNode.nodeValue
            Case "TST_CD"
                giIndex = giIndex + 1
                ReDim Preserve gOrder_List(giIndex)
                ReDim Preserve gResultExamCode(giIndex)
                gOrder_Select.ok = giIndex + 1
                gOrder_List(giIndex).TST_CD = xNode.nodeValue
                
                gResultExamCode(giIndex) = xNode.nodeValue
                
                If Trim(gReceCode) = "" Then
                    gReceCode = "'" & xNode.nodeValue & "'"
                Else
                    gReceCode = gReceCode & ", '" & xNode.nodeValue & "'"
                End If
                
'            Case "WRK_UNT":  gOrder_List(giIndex).WRK_UNT = xNode.nodeValue
'            Case "PT_NM":    gOrder_List(giIndex).PT_NM = xNode.nodeValue
            End Select
            
        End If

        If xNode.hasChildNodes Then
            Display_Order_Parsing xNode.childNodes, Indent
        End If
    Next xNode
End Sub
'Public Sub Display_Order_Parsing(ByRef Nodes As MSXML.IXMLDOMNodeList, _
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
''                giIndex = giIndex + 1
''                ReDim Preserve gOrder_List(giIndex)
''                ReDim Preserve gResultExamCode(giIndex)
'
'                gOrder_List(giIndex).PT_NO = xNode.nodeValue
'            Case "ACPT_DTE": gOrder_List(giIndex).ACPT_DTE = xNode.nodeValue
'            Case "ACPT_NO":  gOrder_List(giIndex).ACPT_NO = xNode.nodeValue
'            Case "TST_CD"
'                giIndex = giIndex + 1
'                ReDim Preserve gOrder_List(giIndex)
'                ReDim Preserve gResultExamCode(giIndex)
'
'                gOrder_List(giIndex).TST_CD = xNode.nodeValue
'
'                gResultExamCode(giIndex) = xNode.nodeValue
'
'                If Trim(gReceCode) = "" Then
'                    gReceCode = "'" & xNode.nodeValue & "'"
'                Else
'                    gReceCode = gReceCode & ", '" & xNode.nodeValue & "'"
'                End If
'
'            Case "WRK_UNT":  gOrder_List(giIndex).WRK_UNT = xNode.nodeValue
'            Case "PT_NM":    gOrder_List(giIndex).PT_NM = xNode.nodeValue
'            End Select
'            gOrder_Select.ok = giIndex
'        End If
'
'        If xNode.hasChildNodes Then
'            Display_Order_Parsing xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub


Public Function Get_QCList(asBarcode As String, asGubun As String) As String
    Dim sRetStr As String
    
    ReDim Preserve gWork_Select(0)
    giIndex = -1
    
    sRetStr = Get_Qry_QCList(asBarcode, asGubun)
    
    SaveXMLFile sRetStr
    
    Dim xDoc As MSXML.DOMDocument

    Set xDoc = New MSXML.DOMDocument

    If xDoc.Load(App.Path & "\order.xml") Then
        ' 문서가 성공적으로 로드되었습니다.
        ' 이제 재미있는 작업을 수행합니다.
        Display_Work_Parsing xDoc.childNodes, 0
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

        Save_Raw_Data strErrText
    End If

    Set xPE = Nothing

    Set xDoc = Nothing
    
    Get_QCList = CStr(gOrder_Select.ok)
End Function



Public Function Get_QCList_Equip(asInputDT As String, asEquip As String) As String
    Dim sRetStr As String
    
    ReDim Preserve gWork_Select(0)
    giIndex = -1
    
    sRetStr = Get_Qry_QCList_Equip(asInputDT, asEquip)
    
    SaveXMLFile sRetStr
    
'    Dim xDoc As MSXML.DOMDocument
'
'    Set xDoc = New MSXML.DOMDocument
'
'    If xDoc.Load(App.Path & "\order.xml") Then
'        ' 문서가 성공적으로 로드되었습니다.
'        ' 이제 재미있는 작업을 수행합니다.
'        Display_Work_Parsing xDoc.childNodes, 0
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
'        Save_Raw_Data strErrText
'    End If
'
'    Set xPE = Nothing
'
'    Set xDoc = Nothing
    
    Get_QCList_Equip = CStr(gOrder_Select.ok)
End Function



'2009.10.01 윤영기
'날짜, 검사코드로 검사리스트 가져오기
'return : 1 => 검사 존재, 0 => 검사 없음
'gOrder_select에 파라미터 저장
Public Function Get_WorkList(asFromDT As String, asToDT As String, asTest As String, asGubun As String) As Integer
    Dim sRetStr As String
    
    ReDim Preserve gWork_Select(0)
    giIndex = -1
    
    sRetStr = Get_Qry_WorkList(asFromDT, asToDT, asTest, asGubun)
    
    SaveXMLFile sRetStr
    
    Dim xDoc As MSXML.DOMDocument
    
    Set xDoc = New MSXML.DOMDocument
    
    If xDoc.Load(App.Path & "\order.xml") Then
        ' 문서가 성공적으로 로드되었습니다.
        ' 이제 재미있는 작업을 수행합니다.
        Display_Work_Parsing xDoc.childNodes, 0
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
        
        Save_Raw_Data strErrText
    End If
    
    Set xPE = Nothing
    
    Set xDoc = Nothing
    
    Get_WorkList = gOrder_Select.ok
End Function

Public Function Get_WorkList1(asFromDT As String, asToDT As String, asTest As String, asGubun As String) As Integer
    Dim sRetStr As String
    
    ReDim Preserve gWork_Select(0)
    giIndex = -1
    
    sRetStr = Get_Qry_WorkList1(asFromDT, asToDT, asTest, asGubun)
    
    SaveXMLFile sRetStr
    
    Dim xDoc As MSXML.DOMDocument
    
    Set xDoc = New MSXML.DOMDocument
    
    If xDoc.Load(App.Path & "\order.xml") Then
        ' 문서가 성공적으로 로드되었습니다.
        ' 이제 재미있는 작업을 수행합니다.
        Display_Work_Parsing1 xDoc.childNodes, 0
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
        
        Save_Raw_Data strErrText
    End If
    
    Set xPE = Nothing

    Set xDoc = Nothing
    
    Get_WorkList1 = gOrder_Select.ok
End Function

'PG_SRL.INTERFACE_S03
'인터페이스 웹서버에서 데이타 가져오기
Public Function Get_Qry_WorkList(asFromDT As String, asToDT As String, asTest As String, asGubun As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    
    strDiv = "PG_SRL.INTERFACE_S15"
    'asSID = "09092251028"
    
    sParam = "<Table>" & _
                      "<QID><![CDATA[PG_SRL.INTERFACE_S15]]></QID>" & _
                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                      "<USERID><![CDATA[LIA]]></USERID>" & _
                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                      "<P0><![CDATA[" & asFromDT & "]]></P0>" & _
                      "<P1><![CDATA[" & asToDT & "]]></P1>" & _
                      "<P2><![CDATA[" & asTest & "]]></P2>" & _
                      "<P3><![CDATA[" & asGubun & "]]></P3>" & _
                      "<P4><![CDATA[" & "" & "]]></P4>" & _
               "</Table>"
                   
    'Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
    
    send = oSOAP.wsLISInterface(strDiv, sParam)
    
    Get_Qry_WorkList = send
    
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
    
    Set oSOAP = Nothing
    
End Function

Public Function Get_Qry_QCList(asBarcode As String, asGubun As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    
    strDiv = "PG_SRL.INTERFACE_S17"
    'asSID = "09092251028"
    
    sParam = "<Table>" & _
                      "<QID><![CDATA[PG_SRL.INTERFACE_S17]]></QID>" & _
                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                      "<USERID><![CDATA[LIA]]></USERID>" & _
                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                      "<P0><![CDATA[" & asGubun & "]]></P0>" & _
                      "<P1><![CDATA[" & asBarcode & "]]></P1>" & _
                      "<P2><![CDATA[]]></P2>" & _
               "</Table>"
                   
    'Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
    
    send = oSOAP.wsLISInterface(strDiv, sParam)
    
    Get_Qry_QCList = send
    
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
    
    Set oSOAP = Nothing
    
End Function

Public Function Get_Qry_QCList_Equip(asInputDT As String, asEquip As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    
    strDiv = "PG_SRL.INTERFACE_S18"
    'asSID = "09092251028"
    
    sParam = "<Table>" & _
                      "<QID><![CDATA[PG_SRL.INTERFACE_S18]]></QID>" & _
                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                      "<USERID><![CDATA[LIA]]></USERID>" & _
                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                      "<P0><![CDATA[" & asInputDT & "]]></P0>" & _
                      "<P1><![CDATA[" & asEquip & "]]></P1>" & _
                      "<P2><![CDATA[]]></P2>" & _
               "</Table>"
                   
    'Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
    
    send = oSOAP.wsLISInterface(strDiv, sParam)
    
    Get_Qry_QCList_Equip = send
    
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
    
    Set oSOAP = Nothing
    
End Function


Public Function Get_Qry_WorkList1(asFromDT As String, asToDT As String, asTest As String, asGubun As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    
    strDiv = "PG_SRL.INTERFACE_S15"
    'asSID = "09092251028"
    
    sParam = "<Table>" & _
                      "<QID><![CDATA[PG_SRL.INTERFACE_S15]]></QID>" & _
                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                      "<USERID><![CDATA[LIA]]></USERID>" & _
                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                      "<P0><![CDATA[" & asFromDT & "]]></P0>" & _
                      "<P1><![CDATA[" & asToDT & "]]></P1>" & _
                      "<P2><![CDATA[" & asTest & "]]></P2>" & _
                      "<P3><![CDATA[" & asGubun & "]]></P3>" & _
                      "<P4><![CDATA[" & "" & "]]></P4>" & _
               "</Table>"
                   
    'Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
    
    send = oSOAP.wsLISInterface(strDiv, sParam)
    
    Get_Qry_WorkList1 = send
    
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

Public Sub Display_QCWork_Parsing(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)

    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
        'Debug.Print xNode.nodeType
        'Debug.Print xNode.nodeType & vbTab & xNode.parentNode.nodeName & " : " & xNode.nodeValue
        If xNode.nodeType = 4 Then
        'If xNode.nodeType = NODE_TEXT Then
        'If xNode.nodeType = NODE_ATTRIBUTE Then
        'If xNode.nodeType = NODE_ELEMENT Then
            'Debug.Print xNode.parentNode.nodeName & " : " & xNode.nodeValue
            
            Select Case xNode.parentNode.nodeName
            Case "INST_DTM":    gQC_Info(giIndex).INST_DTM = xNode.nodeValue
                giIndex = giIndex + 1
                ReDim Preserve gQC_Info(giIndex)
            Case "LOT_NO":    gQC_Info(giIndex).LOT_NO = xNode.nodeValue
            Case "TST_CD"
                gQC_Info(giIndex).TST_CD = xNode.nodeValue
                If gReceCode = "" Then
                    gReceCode = "'" & xNode.nodeValue & "'"
                Else
                    gReceCode = gReceCode & ", '" & xNode.nodeValue & "'"
                End If
            Case "EQUIP_CD":  gQC_Info(giIndex).equip_cd = xNode.nodeValue
            Case "CTRL_CD":  gQC_Info(giIndex).CTRL_CD = xNode.nodeValue
            Case "LOT_NO1":   gQC_Info(giIndex).LOT_NO1 = xNode.nodeValue
            Case "BARCODE_CD": gQC_Info(giIndex).BARCODE_CD = xNode.nodeValue
            Case "USE_YN":    gQC_Info(giIndex).USE_YN = xNode.nodeValue
            
            End Select
            gQC_Info(giIndex).ok = 1
        End If

        If xNode.hasChildNodes Then

            
            Display_QCWork_Parsing xNode.childNodes, Indent
        End If
    Next xNode
End Sub

'XML File Parsing
Public Sub Display_Work_Parsing(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)

    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
        'Debug.Print xNode.nodeType
        'Debug.Print xNode.nodeType & vbTab & xNode.parentNode.nodeName & " : " & xNode.nodeValue
        If xNode.nodeType = 4 Then
        'If xNode.nodeType = NODE_TEXT Then
        'If xNode.nodeType = NODE_ATTRIBUTE Then
        'If xNode.nodeType = NODE_ELEMENT Then
            'Debug.Print xNode.parentNode.nodeName & " : " & xNode.nodeValue
            
            Select Case xNode.parentNode.nodeName
            Case "PT_NO"
                giIndex = giIndex + 1
                ReDim Preserve gWork_Select(giIndex)
                gWork_Select(giIndex).PT_NO = xNode.nodeValue
            Case "PT_NM":    gWork_Select(giIndex).PT_NM = xNode.nodeValue
            Case "SPC_NO"
                gWork_Select(giIndex).SPC_NO = xNode.nodeValue
            Case "TST_DTE":  gWork_Select(giIndex).TST_DTE = xNode.nodeValue
            Case "SEX":  gWork_Select(giIndex).Sex = xNode.nodeValue
            Case "AGE":  gWork_Select(giIndex).Age = xNode.nodeValue
            Case "ACPT_NO":  gWork_Select(giIndex).ACPT_NO = xNode.nodeValue
            Case "TST_CD":   gWork_Select(giIndex).TST_CD = xNode.nodeValue
            Case "TST_STAT": gWork_Select(giIndex).TST_STAT = xNode.nodeValue
            Case "WD_NO":    gWork_Select(giIndex).WD_NO = xNode.nodeValue
            Case "SPC_NM":   gWork_Select(giIndex).SPC_NM = xNode.nodeValue
            
            End Select
            gOrder_Select.ok = 1
        End If

        If xNode.hasChildNodes Then

            
            Display_Work_Parsing xNode.childNodes, Indent
        End If
    Next xNode
End Sub

'XML File Parsing
Public Sub Display_Work_Parsing1(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)

    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
        'Debug.Print xNode.nodeType
        Debug.Print xNode.nodeType & vbTab & xNode.parentNode.nodeName & " : " & xNode.nodeValue
        If xNode.nodeType = 4 Then
        'If xNode.nodeType = NODE_TEXT Then
        'If xNode.nodeType = NODE_ATTRIBUTE Then
        'If xNode.nodeType = NODE_ELEMENT Then
            Debug.Print xNode.parentNode.nodeName & " : " & xNode.nodeValue
            
            Select Case xNode.parentNode.nodeName
            Case "PT_NO":    gWork_Select(giIndex).PT_NO = xNode.nodeValue
            Case "PT_NM":    gWork_Select(giIndex).PT_NM = xNode.nodeValue
            Case "SPC_NO":   gWork_Select(giIndex).SPC_NO = xNode.nodeValue
            Case "TST_DTE":  gWork_Select(giIndex).TST_DTE = xNode.nodeValue
            Case "ACPT_NO":  gWork_Select(giIndex).ACPT_NO = xNode.nodeValue
            Case "TST_CD":   gWork_Select(giIndex).TST_CD = xNode.nodeValue
            Case "TST_STAT": gWork_Select(giIndex).TST_STAT = xNode.nodeValue
            Case "WD_NO":    gWork_Select(giIndex).WD_NO = xNode.nodeValue
            Case "SPC_NO":   gWork_Select(giIndex).SPC_NO = xNode.nodeValue
            Case "SPC_NM":   gWork_Select(giIndex).SPC_NM = xNode.nodeValue
            
            End Select
            gOrder_Select.ok = 1
        End If

        If xNode.hasChildNodes Then
            giIndex = giIndex + 1
            ReDim Preserve gWork_Select(giIndex)
            
            Display_Work_Parsing1 xNode.childNodes, Indent
        End If
    Next xNode
End Sub


'2009.10.01 윤영기
'검체번호로 환자정보 가져오기
'return : 1 => 검사 존재, 0 => 검사 없음
'gOrder_select에 파라미터 저장
Public Function Get_PatInfo(asSID) As Integer
    Dim sRetStr As String
    
    gOrder_Select.ok = 0
    
    sRetStr = Get_Qry_PatInfo(asSID)
    
    SaveXMLFile sRetStr
    
    Dim xDoc As MSXML.DOMDocument
    
    Set xDoc = New MSXML.DOMDocument
    
    If xDoc.Load(App.Path & "\order.xml") Then
        ' 문서가 성공적으로 로드되었습니다.
        ' 이제 재미있는 작업을 수행합니다.
        Display_PatInfo_Parsing xDoc.childNodes, 0
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
        
        Save_Raw_Data strErrText
    End If
    
    Set xPE = Nothing

    Set xDoc = Nothing
    
    Get_PatInfo = gOrder_Select.ok
End Function

'PG_SRL.INTERFACE_S06
'인터페이스 웹서버에서 데이타 가져오기
Public Function Get_Qry_PatInfo(ByVal asSID As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    
    strDiv = "PG_SRL.INTERFACE_S06"
    'asSID = "09092251028"
    
    sParam = "<Table>" & _
                      "<QID><![CDATA[PG_SRL.INTERFACE_S06]]></QID>" & _
                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                      "<USERID><![CDATA[LIA]]></USERID>" & _
                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                      "<P0><![CDATA[" & asSID & "]]></P0>" & _
                      "<P1><![CDATA[" & "" & "]]></P1>" & _
               "</Table>"
                   
'    Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
    
    send = oSOAP.wsLISInterface(strDiv, sParam)
    
    Get_Qry_PatInfo = send
    
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

'XML File Parsing
Public Sub Display_PatInfo_Parsing(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
        If xNode.nodeType = 4 Then
        'If xNode.nodeType = NODE_TEXT Then
        'If xNode.nodeType = NODE_ATTRIBUTE Then
        'If xNode.nodeType = NODE_ELEMENT Then
            Select Case xNode.parentNode.nodeName
            Case "PTNO"
                gPatient_Info.PTNO = xNode.nodeValue
                
                gOrder_Select.ok = 1
            Case "PATNAME":  gPatient_Info.PATNAME = xNode.nodeValue
            Case "SEX":      gPatient_Info.Sex = xNode.nodeValue
            Case "AGE":      gPatient_Info.Age = xNode.nodeValue
            Case "WD_NO":    gPatient_Info.WD_NO = xNode.nodeValue
            Case "SPC_CD":   gPatient_Info.SPC_CD = xNode.nodeValue
            Case "SPC_NM":   gPatient_Info.SPC_NM = xNode.nodeValue
            Case "ACPT_NO":  gPatient_Info.ACPT_NO = xNode.nodeValue
            Case "ACPT_DTM": gPatient_Info.ACPT_DTM = xNode.nodeValue
            Case "TST_STAT": gPatient_Info.TST_STAT = xNode.nodeValue
            End Select
            
        End If

        If xNode.hasChildNodes Then
            Display_PatInfo_Parsing xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Function Online_QCResult(ByVal asRes As String, _
                              ByVal asEquip As String, _
                              ByVal asLotno As String, _
                              ByVal asExam As String, _
                              ByVal asSpcno As String, _
                              ByVal asInstDTM As String, _
                              ByVal asGubun As String) As String
    
    
    Dim sRetStr As String
    
    
    Online_QCResult = ""
    
    gOnline_Ret = ""
    
    sRetStr = Online_QCResult_Qry(asRes, asEquip, asLotno, asExam, asSpcno, asInstDTM, asGubun)
    
    SaveXMLFile sRetStr
    
    Dim xDoc As MSXML.DOMDocument
    
    Set xDoc = New MSXML.DOMDocument
    
    If xDoc.Load(App.Path & "\order_online.xml") Then
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
        
        Save_Raw_Data strErrText
    End If
    
    Set xPE = Nothing

    Set xDoc = Nothing
    
    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_QCResult = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If
    
End Function

Public Function Online_Result(ByVal asSpcno As String, _
                              ByVal asExam As String, _
                              ByVal asRes As String, _
                              ByVal asEquip As String, _
                              ByVal asCount As String) As String
    
    
    Dim sRetStr As String
    
    
    Online_Result = ""
    
    gOnline_Ret = ""
    
    sRetStr = Online_Result_Qry(asSpcno, asExam, asRes, asEquip, asCount)
    
    SaveXMLFile sRetStr
    
    Dim xDoc As MSXML.DOMDocument
    
    Set xDoc = New MSXML.DOMDocument
    
    If xDoc.Load(App.Path & "\order_online.xml") Then
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
        
        Save_Raw_Data strErrText
    End If
    
    Set xPE = Nothing

    Set xDoc = Nothing
    
    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_Result = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If
    
End Function

Public Function Online_QCResult_Qry(ByVal asRes As String, _
                              ByVal asEquip As String, _
                              ByVal asLotno As String, _
                              ByVal asExam As String, _
                              ByVal asSpcno As String, _
                              ByVal asInstDTM As String, _
                              ByVal asGubun As String) As String
                              
'                              asRes, asEquip, asLotno, asExam, asSpcno, asInstDTM, asGubun
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    
    strDiv = "PG_SRL.INTERFACE_U04"
    'asSID = "09092251028"
    
    sParam = "<Table>" & _
                      "<QID><![CDATA[PG_SRL.INTERFACE_U04]]></QID>" & _
                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                      "<USERID><![CDATA[LIA]]></USERID>" & _
                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                      "<P0><![CDATA[" & asRes & "]]></P0>" & _
                      "<P1><![CDATA[" & asEquip & "]]></P1>" & _
                      "<P2><![CDATA[" & asLotno & "]]></P2>" & _
                      "<P3><![CDATA[" & asExam & "]]></P3>" & _
                      "<P4><![CDATA[" & asSpcno & "]]></P4>" & _
                      "<P5><![CDATA[" & asInstDTM & "]]></P5>" & _
                      "<P6><![CDATA[" & asGubun & "]]></P6>" & _
               "</Table>"
                   
'    Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
    
    send = oSOAP.wsLISInterface(strDiv, sParam)
    
    Online_QCResult_Qry = send
    
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


Public Function Online_Result_Qry(ByVal asSpcno As String, _
                              ByVal asExam As String, _
                              ByVal asRes As String, _
                              ByVal asEquip As String, _
                              ByVal asCount As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    
    strDiv = "PG_SRL.INTERFACE_I01"
    'asSID = "09092251028"
    
    sParam = "<Table>" & _
                      "<QID><![CDATA[PG_SRL.INTERFACE_I01]]></QID>" & _
                      "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                      "<USERID><![CDATA[LIA]]></USERID>" & _
                      "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                      "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                      "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                      "<P1><![CDATA[" & asExam & "]]></P1>" & _
                      "<P2><![CDATA[" & asRes & "]]></P2>" & _
                      "<P3><![CDATA[]]></P3>" & _
                      "<P4><![CDATA[" & asEquip & "]]></P4>" & _
                      "<P5><![CDATA[]]></P5>" & _
                      "<P6><![CDATA[" & asCount & "]]></P6>" & _
                      "<P7><![CDATA[]]></P7>" & _
                      "<P8><![CDATA[]]></P8>" & _
                      "<P9><![CDATA[]]></P9>" & _
                      "<P10><![CDATA[]]></P10>" & _
               "</Table>"
                   
'    Save_Raw_Data "New_SelectOrder Param : " & vbCrLf & sParam
    
    send = oSOAP.wsLISInterface(strDiv, sParam)
    
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


'XML File Parsing
Public Sub Display_Online_Parsing(ByRef Nodes As MSXML.IXMLDOMNodeList, _
    ByVal Indent As Integer)
    
    Dim xNode As MSXML.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Ret = gOnline_Ret & xNode.nodeValue & vbTab
        End If

        If xNode.hasChildNodes Then
            Display_Online_Parsing xNode.childNodes, Indent
        End If
    Next xNode
End Sub

