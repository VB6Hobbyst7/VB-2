Attribute VB_Name = "Module1"
Option Explicit

Public gOnline_Ret As String
Public gOnline_Test As String
Public gServerPath As String
Public gIFUser As String
Public giIndex  As Long
Public gOrderExam As String

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

'''        SaveData strErrText
    End If

    Set xPE = Nothing

    Set xDoc = Nothing

    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_Result = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If

End Function

Public Function Online_Result_Qry(ByVal asParam As String) As String

  Dim o As New XMLHTTPRequest
  Dim s As String
  Dim txtResponseHeaders  As String
  Dim txtResponse  As String
  
On Error GoTo err_handler

  s = s & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
  s = s & "<SOAP-ENV:Envelope" & vbCrLf
  s = s & "SOAP-ENV:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/""" & vbCrLf
  s = s & "xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/""" & vbCrLf
  s = s & "xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/""" & vbCrLf
  s = s & "xmlns:ns0=""capeconnect:GlobalWeather:GlobalWeather""" & vbCrLf
  s = s & "xmlns:xsd=""http://www.w3.org/2001/XMLSchema""" & vbCrLf
  s = s & "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & vbCrLf
  s = s & "<SOAP-ENV:Body>" & vbCrLf
  s = s & "<ns0:getWeatherReport>" & vbCrLf
  s = s & "<code xsi:type=""xsd:string"">CYVR</code>" & vbCrLf
  s = s & "</ns0:getWeatherReport>" & vbCrLf
  s = s & "</SOAP-ENV:Body>" & vbCrLf
  s = s & "</SOAP-ENV:Envelope>" & vbCrLf
  
    s = "<?xml version='1.0' encoding='UTF-8'?>"
    s = s & vbCrLf & "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>"
    s = s & vbCrLf & "<soapenv:Body>"
    s = s & vbCrLf & "<registSpcmRcpn xmlns='http://svc.poct.ws.nhimc/'>"
    s = s & vbCrLf & "<arg0 xmlns=''>" & "1606300050</arg0>"
    s = s & vbCrLf & "</registSpcmRcpn>"
    s = s & vbCrLf & "</soapenv:Body>"
    s = s & vbCrLf & "</soapenv:Envelope>" & vbCrLf
      
'  o.open "POST", "http://192.168.5.105:8800/service/PoctService?wsml", False ' "http://live.capescience.com:80/ccx/GlobalWeather", False
  o.open "POST", "http://192.168.5.105:8800/service/PoctService?wsdl", False ' "http://live.capescience.com:80/ccx/GlobalWeather", False
  o.setRequestHeader "Content-Type", "text/xml"
'  o.setRequestHeader "Connection", "close"
  o.setRequestHeader "Connection", "PoctService"
  o.setRequestHeader "SOAPAction", ""
 ' o.send s
   o.send "1607000010"
  txtResponseHeaders = o.getAllResponseHeaders
  txtResponse = o.responseText

err_handler:
  If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description



    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam
    Dim strXML As String
    
    strXML = "<?xml version='1.0' encoding='UTF-8'?>"
    strXML = strXML & vbCrLf & "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>"
    strXML = strXML & vbCrLf & "<soapenv:Body>"
    strXML = strXML & vbCrLf & "<registSpcmRcpn xmlns='http://svc.poct.ws.nhimc/'>"
    strXML = strXML & vbCrLf & "<arg0 xmlns=''>" & "1585000010</arg0>"
    strXML = strXML & vbCrLf & "</registSpcmRcpn>"
    strXML = strXML & vbCrLf & "</soapenv:Body>"
    strXML = strXML & vbCrLf & "</soapenv:Envelope>"
    
    
    
    On Error GoTo ErrHandle

    Set oSOAP = New MSSOAPLib30.SoapClient30


   sParam = "http://192.168.5.105:8800/service/PoctService?wsdl"
    
Dim Client As New SoapClient30


    Client.ClientProperty("ServerHTTPRequest") = True
    Client.MSSoapInit sParam
    Client.ConnectorProperty("EndPointURL") = "http://192.168.5.105:8800/service/PoctService?wsml"
    Client.ConnectorProperty("AuthUser") = "test"
    Client.ConnectorProperty("AuthPassword") = "test"
    

send = Client.PoctService(strXML)

'    Call oSOAP.ConnectorProperty("PoctService")
    
    strDiv = "PoctService"

'    Call oSOAP.MSSoapInit("http://192.168.5.105:8800/service/PoctService?wsdl", strDiv, "PoctPort", "http://192.168.5.105:8800/service/PoctService?wsml")
    Call oSOAP.MSSoapInit("http://192.168.5.105:8800/service/PoctService?wsdl")

'http://192.168.5.105:8800/service/PoctService?wsml

    oSOAP.ClientProperty("ServerHTTPRequest") = True
    strDiv = "PoctService"
    'oSOAP.MSSoapInit Form1.txtServerPath.Text, "PoctService", "PoctPort", "http://192.168.5.105:8800/service/PoctService?wsml"
    'oSOAP.MSSoapInit2 Form1.txtServerPath.Text ', "http://192.168.5.105:8800/service/PoctService?wsml",
    '"http://svc.poct.ws.nhimc/PoctService", "http://svc.poct.ws.nhimc/PoctPort", "http://svc.poct.ws.nhimc/"

    ' 서비스명 확인해서 테스트 요망 ===================================


    'strDiv = "http://192.168.5.105:8800/service/PoctService"
    sParam = asParam



    'send = oSOAP.wsLISInterhhhface("Poct.registSpcmRcpn", "1111")
'    send = oSOAP.wsLISInterface(Val(strDiv), sParam)
    'send = oSOAP.wsPoctService(Val(strDiv), sParam)
    'send = oSOAP.registSpcmRcpn("123345")
    send = oSOAP.registSpcmRcpn(strXML)



    ' ==================================================================
    Online_Result_Qry = send
    Set oSOAP = Nothing
    DoEvents
    Exit Function

ErrHandle:
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & "/" & vbCrLf & oSOAP.Detail & vbCrLf

    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If



End Function

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
