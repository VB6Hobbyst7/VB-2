Attribute VB_Name = "Kbnuweb"

Option Explicit

Type Order_Select
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
    DEPT_CD     As String
    WD_NO       As String
    SPC_NM      As String
    SPC_CD      As String
    SPC_LAST    As String
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
Public gQC_Select As QC_Info
Public gQC_Info() As QC_Info

Public gOnline_Ret As String
Public gReceCode As String

'추가 변수 start==========================================
Type Worker_Info
    WK_ID       As String
    WK_PW       As String
    WK_NM       As String
    ok          As Integer
End Type
Public gWorker_Info As Worker_Info

Public Function kbnu_Server_Connect(ByVal asEquip As String) As String
'  Dim XMLRequest            As New XMLHTTPRequest
  Dim txtSendXML               As String
  Dim txtResponseHeaders    As String
  Dim txtResponse           As String
  
  
On Error GoTo err_handler

    Dim lsConnectReq As String
    Dim lsEquipCDReq As String
    Dim lsBarcodeReq As String
    
    lsConnectReq = "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00113&business_id=lis&bcno=&testcd=&eqmtcd=" & asEquip & "&instcd=032&"
    'lsEquipCDReq = "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00103&business_id=lis&refgbn=2&instcd=032&eqmtcd=H05&"
    
    'txtServerPath = lsConnectReq
    
    XMLRequest.Open "POST", lsConnectReq, False
    'XMLRequest.setRequestHeader "Content-Type", "text/xml"
    'XMLRequest.setRequestHeader "Connection", "PoctService"
    'XMLRequest.setRequestHeader "SOAPAction", ""
    XMLRequest.send ""
    
    'XMLRequest.send "1607000010"
    
    
    'txtResponseHeaders = XMLRequest.getAllResponseHeaders
    txtResponse = XMLRequest.responseText
    
    kbnu_Server_Connect = txtResponse
    
    Save_Xml_Data txtResponse, "connect"
'    txtXML.Text = txtResponse
'    txtResponse = Replace(txtResponse, ">", ">" & vbNewLine)
'    txtXMLedit.Text = txtResponse
    
    Exit Function
    
err_handler:
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Function

Public Function kbnu_Equipcode(ByVal asEquip As String) As String
  Dim XMLRequest            As New XMLHTTPRequest
  Dim txtSendXML               As String
  Dim txtResponseHeaders    As String
  Dim txtResponse           As String
  
  
On Error GoTo err_handler

    Dim lsConnectReq As String
    Dim lsEquipCDReq As String
    Dim lsBarcodeReq As String
    
    'http://his031.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00103&business_id=lis&refgbn=2&instcd=031&eqmtcd=H05&
    lsEquipCDReq = "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00103&business_id=lis&refgbn=2&instcd=032&eqmtcd=" & asEquip & "&"
    Save_Raw_Data lsEquipCDReq
    'txtServerPath = lsEquipCDReq
    
    XMLRequest.Open "POST", lsEquipCDReq, False
    'XMLRequest.setRequestHeader "Content-Type", "text/xml"
    'XMLRequest.setRequestHeader "Connection", "PoctService"
    'XMLRequest.setRequestHeader "SOAPAction", ""
    XMLRequest.send ""
    
    'XMLRequest.send "1607000010"
    
    
    'txtResponseHeaders = XMLRequest.getAllResponseHeaders
    txtResponse = XMLRequest.responseText
    kbnu_Equipcode = txtResponse
    
    Save_Raw_Data txtResponse
    
    Save_Xml_Data txtResponse, "equipcode"
    
'    txtXML.Text = txtResponse
'
'    txtResponse = Replace(txtResponse, ">", ">" & vbNewLine)
'    txtXMLedit.Text = txtResponse
    
    Exit Function
    
err_handler:
  If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Public Function kbnu_Order_Request(ByVal asBarcode As String, ByVal asEquip As String) As String
  Dim XMLRequest            As New XMLHTTPRequest
  Dim txtSendXML               As String
  Dim txtResponseHeaders    As String
  Dim txtResponse           As String
  
  
On Error GoTo err_handler

    Dim lsReqMsg As String
    
    'http://his031.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00101&business_id=lis&bcno=E145Z0050&instcd=031&eqmtcd=H05&
    lsReqMsg = "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00101&business_id=lis&bcno=" & asBarcode & "&instcd=032&eqmtcd=P01&"

        
    Save_Raw_Data lsReqMsg
    'txtServerPath = lsReqMsg
    
    XMLRequest.Open "POST", lsReqMsg, False
    'XMLRequest.setRequestHeader "Content-Type", "text/xml"
    'XMLRequest.setRequestHeader "Connection", "PoctService"
    'XMLRequest.setRequestHeader "SOAPAction", ""
    XMLRequest.send ""
    
    'XMLRequest.send "1607000010"
    
    
    'txtResponseHeaders = XMLRequest.getAllResponseHeaders
    txtResponse = XMLRequest.responseText
    Save_Raw_Data txtResponse
    
    txtResponse = Replace(txtResponse, "><", ">" & vbLf & "<")
    txtResponse = Replace(txtResponse, "▦", vbTab)
    
    Save_Raw_Data txtResponse
    
    kbnu_Order_Request = txtResponse
    
    Save_Xml_Data txtResponse, "order"
    

    gOrder_Select.ok = 0
    
    giIndex = -1
    ReDim gOrder_List(0)
    
    'sRetStr = Get_Qry_OrderList(asSID)
    
    'SaveXMLFile sRetStr
    'Save_Xml_Data sRetStr, "online_order"
    
    Dim lsLine As String
    Dim lsGubun As String
    Dim lsData As String
    Dim FilNum
    Dim i, j, cnt As Integer
    
    FilNum = FreeFile
    
    lsLine = txtResponse
    
    i = InStr(1, lsLine, vbLf)
    Do While i > 0
        lsData = Left(lsLine, i - 1)
        lsLine = Mid(lsLine, i + 1)
        
        Select Case lsData
            Case "<spcacptdt>"
                lsGubun = lsData
            Case "<acptflag>"
                lsGubun = lsData
            Case "<bcno>"
                lsGubun = lsData
            Case "<pid>"
                lsGubun = lsData
            Case "<patnm>"
                lsGubun = lsData
            Case "<sexage>"
                lsGubun = lsData
            Case "<workno>"
                lsGubun = lsData
            Case "<tsectnm>"
                lsGubun = lsData
            Case "<ifreqcdlist>"
                lsGubun = lsData
            Case "<tclscdlist>"
                lsGubun = lsData
            Case "<urinextrvol>"
                lsGubun = lsData
            Case "<retestyn>"
                lsGubun = lsData
            Case "<rsltstat>"
                lsGubun = lsData
            Case "<spccd>"
                lsGubun = lsData
            Case "<spcnm>"
                lsGubun = lsData
            Case Else
                'lsGubun = ""
        End Select
        
        If InStr(1, lsData, "<![CDATA") > 0 Then
            lsData = Replace(lsData, "<![CDATA[", "")
            lsData = Replace(lsData, "[", "")
            lsData = Replace(lsData, "]", "")
            lsData = Replace(lsData, ">", "")
            lsData = Replace(lsData, "▦", vbTab)
            
            Select Case lsGubun
                Case "<spcacptdt>"
                    gOrder_Select.ACPT_DTE = lsData
                Case "<acptflag>"
                    '
                Case "<bcno>"
                    gOrder_Select.SPC_NO = lsData
                    gOrder_Select.ok = 1
                Case "<pid>"
                    gOrder_Select.PT_NO = lsData
                Case "<patnm>"
                    gOrder_Select.PT_NM = lsData
                Case "<sexage>"
                    gOrder_Select.Sex = lsData
                Case "<workno>"
                    gOrder_Select.ACPT_NO = lsData
                Case "<tsectnm>"
                    lsGubun = lsData
                Case "<ifreqcdlist>"
                    lsGubun = lsData
                Case "<tclscdlist>"
                    '▦
                    giIndex = 0
                    j = InStr(1, lsData, vbTab)
                    Do While j > 0
                        giIndex = giIndex + 1
                        
                        If UBound(gOrder_List) < giIndex Then
                            ReDim Preserve gOrder_List(giIndex)
                        End If
                        
                        gOrder_List(giIndex).TST_CD = Left(lsData, j - 1)
                        gOrder_List(giIndex).ok = 1
                        
                        j = InStr(j + 1, lsData, vbTab)
                    Loop
                Case "<urinextrvol>"
                    lsGubun = lsData
                Case "<retestyn>"
                    giIndex = 0
                    j = InStr(1, lsData, vbTab)
                    Do While j > 0
                        giIndex = giIndex + 1
                        
                        If UBound(gOrder_List) < giIndex Then
                            ReDim Preserve gOrder_List(giIndex)
                        End If
                        
                        gOrder_List(giIndex).TST_DTE = Left(lsData, j - 1)
                        
                        
                        j = InStr(j + 1, lsData, vbTab)
                    Loop
                Case "<rsltstat>"
                    giIndex = 0
                    j = InStr(1, lsData, vbTab)
                    Do While j > 0
                        giIndex = giIndex + 1
                        
                        If UBound(gOrder_List) < giIndex Then
                            ReDim Preserve gOrder_List(giIndex)
                        End If
                        
                        gOrder_List(giIndex).TST_STAT = Left(lsData, j - 1)
                        
                        
                        j = InStr(j + 1, lsData, vbTab)
                    Loop
                Case "<spccd>"
                    gOrder_Select.SPC_CD = lsData
                Case "<spcnm>"
                    gOrder_Select.SPC_NM = lsData
                Case Else
                    
            End Select
            lsGubun = ""
        End If
        
        i = InStr(1, lsLine, vbLf)
    Loop

    
    Exit Function
    
err_handler:
  If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
  
End Function


Public Function kbnu_sendresult(ByVal asBarcode As String, ByVal asUID As String, ByVal asEquip As String, ByVal asResult As String) As String
  Dim XMLRequest            As New XMLHTTPRequest
  Dim txtSendXML               As String
  Dim txtResponseHeaders    As String
  Dim txtResponse           As String
  
  
On Error GoTo err_handler

    Dim lsConnectReq As String
    Dim lsEquipCDReq As String
    Dim lsBarcodeReq As String
    
    '               http://his031.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TXLII00101&business_id=lis&ex_interface=93367|031&bcno=E145Z0050&result=LHC1020190.420101228LHC102021.0620101228LHC1020411.620101228LHC1030229.920101228&instcd=031&eqmtcd=H05&userid=93367&paste=Y&
    lsEquipCDReq = "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TXLII00101&business_id=lis&ex_interface=" & asUID & "|032&bcno=" & asBarcode & "&result=" & asResult & "&instcd=032&eqmtcd=" & asEquip & "&userid=" & asUID & "&paste=Y&"
    
    Save_Raw_Data lsEquipCDReq
    
    'txtServerPath = lsEquipCDReq
    
    XMLRequest.Open "POST", lsEquipCDReq, False
    'XMLRequest.setRequestHeader "Content-Type", "text/xml"
    'XMLRequest.setRequestHeader "Connection", "PoctService"
    'XMLRequest.setRequestHeader "SOAPAction", ""
    XMLRequest.send ""
    
    'XMLRequest.send "1607000010"
    
    
    'txtResponseHeaders = XMLRequest.getAllResponseHeaders
    txtResponse = XMLRequest.responseText
    
    Save_Raw_Data txtResponse
    
    kbnu_sendresult = txtResponse
    
    Save_Xml_Data txtResponse, "equipcode"
    
'    txtXML.Text = txtResponse
'
'    txtResponse = Replace(txtResponse, ">", ">" & vbNewLine)
'    txtXMLedit.Text = txtResponse
    
    Exit Function
    
err_handler:
  If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Public Function Get_Order(asSID) As Integer
    Dim sRetStr As String
    
    
    
    Dim xDoc As MSXML.DOMDocument
    
    Set xDoc = New MSXML.DOMDocument
    
    If xDoc.Load(App.Path & "\Res" & "\order.xml") Then
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
        
        SaveData strErrText
    End If
    
    Set xPE = Nothing

    Set xDoc = Nothing
    
    Get_Order = gOrder_Select.ok
End Function

'Public Sub Display_Order_Parsing(ByRef Nodes As MSXML.IXMLDOMNodeList, ByVal Indent As Integer)
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
'                giIndex = giIndex + 1
'                ReDim Preserve gOrder_List(giIndex)
'
'                gOrder_List(giIndex).PT_NO = xNode.nodeValue
'            Case "ACPT_DTE": gOrder_List(giIndex).ACPT_DTE = xNode.nodeValue
'            Case "ACPT_NO":  gOrder_List(giIndex).ACPT_NO = xNode.nodeValue
'            Case "TST_CD":   gOrder_List(giIndex).TST_CD = xNode.nodeValue
'            Case "WRK_UNT":  gOrder_List(giIndex).WRK_UNT = xNode.nodeValue
'            Case "PT_NM":    gOrder_List(giIndex).PT_NM = xNode.nodeValue
'            End Select
'            gOrder_Select.ok = 1
'        End If
'
'        If xNode.hasChildNodes Then
'            Display_Order_Parsing xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub

Public Function Get_Qry_OrderList(ByVal asSID As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    oSOAP.MSSoapInit "http://his032.knu.ac.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00101&business_id=lis&bcno=" & asSID & "instcd=032&eqmtcd=P01&"
    strDiv = "PG_SRL.INTERFACE_S03"
    'asSID = "09092251028"
    
    sParam = "<Table>" & _
                      "<QID><![CDATA[PG_SRL.INTERFACE_S03]]></QID>" & _
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

Public Function Get_Qry_WorkList(asFromDT As String, asToDT As String, asTest As String, asGubun As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
'    oSOAP.MSSoapInit "http://interface.gnuh.co.kr/WEBSERVICE/INTERFACE/LisInterface.asmx?wsdl"
    
'    oSOAP
    strDiv = "PG_SRL.INTERFACE_S15"

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

'Public Sub Display_Work_Parsing(ByRef Nodes As MSXML.IXMLDOMNodeList, _
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
'            Case "PT_NO"
'                giIndex = giIndex + 1
'                ReDim Preserve gWork_Select(giIndex)
'                gWork_Select(giIndex).PT_NO = xNode.nodeValue
'            Case "PT_NM":    gWork_Select(giIndex).PT_NM = xNode.nodeValue
'            Case "SPC_NO"
'                gWork_Select(giIndex).SPC_NO = xNode.nodeValue
'            Case "TST_DTE":  gWork_Select(giIndex).TST_DTE = xNode.nodeValue
'            Case "SEX":  gWork_Select(giIndex).Sex = xNode.nodeValue
'            Case "AGE":  gWork_Select(giIndex).Age = xNode.nodeValue
'            Case "ACPT_NO":  gWork_Select(giIndex).ACPT_NO = xNode.nodeValue
'            Case "TST_CD":   gWork_Select(giIndex).TST_CD = xNode.nodeValue
'            Case "TST_STAT": gWork_Select(giIndex).TST_STAT = xNode.nodeValue
'            Case "DEPT_CD":    gWork_Select(giIndex).DEPT_CD = xNode.nodeValue
'            Case "WD_NO":    gWork_Select(giIndex).WD_NO = xNode.nodeValue
'            Case "SPC_NM":   gWork_Select(giIndex).SPC_NM = xNode.nodeValue
'            Case "ACPT_DTM":  gWork_Select(giIndex).ACPT_DTE = xNode.nodeValue
'
'            End Select
'            gOrder_Select.ok = 1
'        End If
'
'        If xNode.hasChildNodes Then
'
'
'            Display_Work_Parsing xNode.childNodes, Indent
'        End If
'    Next xNode
'End Sub

Public Sub Save_Xml_Data(argSQL As String, argFileName As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\" & "Res", vbDirectory) <> "Res" Then
        MkDir (App.Path & "\" & "Res")
    End If
    
'    sFileName = Format(CDate(frmMain.txtToday.Text), "yyyymmdd")
    sFileName = argFileName
    If Dir(App.Path & "\" & "Res" & "\" & sFileName & ".xml") <> "" Then
        Kill App.Path & "\" & "Res" & "\" & sFileName & ".xml"
    End If
    
    Open App.Path & "\" & "Res" & "\" & sFileName & ".xml" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

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
