Attribute VB_Name = "modPubHealth"
Option Explicit

Public Function Get_WorkList() As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim send
    Dim sParam
    Dim sRet
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl"
    'oSOAP.MSSoapInit gAddr
    
    
    sParam = "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & Chr(13)
    sParam = sParam & "PID|||^" & gHPEquip & "^" & gUID & "^DefaultDomain^PI" & Chr(13)
    sParam = sParam & "PV1||E|" & gHPID & Chr(13)
    sParam = sParam & "OBR|1||||||1" & Chr(13)
    sParam = Chr(11) & sParam & Chr(12) & Chr(13)
    'Debug.Print sParam
    
'    Save_Raw_Data "Worklist Param : " & vbCrLf & sParam
    
    sParam = makeB64(sParam)
    
    'MsgBox oSOAP.detail
    
    send = oSOAP.MdbOrderList(sParam)
    
    send = makeUB64(send)
    
'    Save_Raw_Data "Worklist Return : " & vbCrLf & send
    
    Get_WorkList = send
    
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

Public Function Get_OrderList(asSID As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
        
    'oSOAP.mssoapinit "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl"
    oSOAP.MSSoapInit gAddr
    
    sParam = "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & Chr(13) & Chr(10)
    sParam = sParam & "PID|||" & asSID & "^" & gHPEquip & "^" & gUID & "^DefaultDomain^PI" & Chr(13) & Chr(10)
    sParam = sParam & "PV1||E|" & gHPID & Chr(13) & Chr(10)
    sParam = sParam & "OBR|1||||||1" & Chr(13) & Chr(10)
    sParam = Chr(11) & sParam & Chr(12) & Chr(13)
    
    sParam = makeB64(sParam)
    
    send = oSOAP.New_SelectOrder(sParam)
    
    send = makeUB64(send)
    
    SetRawData "New_SelectOrder Return : " & vbCrLf & send
    
    Get_OrderList = send
    
    'MsgBox send
    
   ' MsgBox send
    
    SetRawData "[Rcv]" & Get_OrderList
    Set oSOAP = Nothing

    DoEvents
    
    Exit Function

ErrHandle:

    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        'SetRawData "[Err]" & Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
        'SetRawData "[Err]" & Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
    
End Function


Public Function SendResult(asParam As String) As Integer
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim send As String
    Dim sParam As String
    
    SendResult = 0
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    'oSOAP.mssoapinit "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl"
    oSOAP.MSSoapInit gAddr
    
    sParam = asParam
    
'    sParam = "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & Chr(13)
'    'sParam = sParam & Chr(11) & "PID|||^CBC^B080100043^DefaultDomain^PI" & Chr(12) & Chr(13)
'    'sParam = sParam & Chr(11) & "PID|||^^B080100043^DefaultDomain^PI" & Chr(12) & Chr(13)
'    'sParam = sParam & Chr(11) & "PID|||^AIDS^B080100043^DefaultDomain^PI" & Chr(12) & Chr(13)
'    sParam = sParam & "PID|||" & asSID & "^C1^" & gUID & "^DefaultDomain^PI" & Chr(13)
'    sParam = sParam & "PV1||E|" & gHPID & Chr(13)
'    sParam = sParam & "OBR|1||||||1" & Chr(13)
'    sParam = Chr(11) & sParam & Chr(12) & Chr(13)
'    'Debug.Print sParam
    
'    Save_Raw_Data "UpdateRst Param : " & vbCrLf & sParam
    
    sParam = makeB64(sParam)
    
    send = oSOAP.UpdateRst(sParam)
    
    SetRawData "UpdateRst Ret : " & vbCrLf & send

    'send = makeUB64(send)
    
    SendResult = CInt(send)
    
    Set oSOAP = Nothing

    
    DoEvents
    
    Exit Function

ErrHandle:
    SendResult = -1
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
End Function

