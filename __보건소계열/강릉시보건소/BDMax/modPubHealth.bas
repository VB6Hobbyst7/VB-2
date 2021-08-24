Attribute VB_Name = "modPubHealth"
Option Explicit

Public Function Get_WorkList() As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim send
    Dim sParam
    Dim sRet
    
    'On Error GoTo ErrHandle
    'MsgBox "1"
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    
    'MsgBox "2"
    
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    'MsgBox gAddr
    'oSOAP.MSSoapInit "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl"
    oSOAP.MSSoapInit gAddr
    
    'MsgBox "3"
    
    
    sParam = "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & Chr(13) & Chr(10)
    sParam = sParam & "PID|||^" & gHPEquip & "^" & gUID & "^DefaultDomain^PI" & Chr(13) & Chr(10)
    sParam = sParam & "PV1||E|" & gHPID & Chr(13) & Chr(10)
    sParam = sParam & "OBR|1||||||1" & Chr(13) & Chr(10)
    sParam = Chr(11) & sParam & Chr(12) & Chr(13)
    'Debug.Print sParam
    
    'MsgBox sParam

    SetRawData "Worklist Param : " & vbCrLf & sParam
    
    sParam = makeB64(sParam)
    
    'MsgBox oSOAP.detail
    
    send = oSOAP.MdbOrderList(sParam)
    
    send = makeUB64(send)
    
    
'    send = ""
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910173|10117020|ÀÏ¾ï¼ö»ê¼öÁ·°ü¼ö|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910174|10117020|ÀÏ¾ï¼ö»ê¼öÁ·°ü¼ö|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910175|10117020|ÀÏ¾ï¼ö»ê¼öÁ·°ü¼ö|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910176|10117020|ÀÏ¾ï¼ö»ê¼öÁ·°ü¼ö|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910181|10117020|ÀÏ¾ï¼ö»ê¼öÁ·°ü¼ö|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910182|10117020|ÀÏ¾ï¼ö»ê¼öÁ·°ü¼ö|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910183|10117020|ÀÏ¾ï¼ö»ê¼öÁ·°ü¼ö|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910184|10117020|ÀÏ¾ï¼ö»ê¼öÁ·°ü¼ö|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910188|10117020|ÀÏ¾ï¼ö»ê¼öÁ·°ü¼ö|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910173|10117021|»ï¾ç´ßÁý¾ç³ä¼Ò½º|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910174|10117021|»ï¾ç´ßÁý¾ç³ä¼Ò½º|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910175|10117021|»ï¾ç´ßÁý¾ç³ä¼Ò½º|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910176|10117021|»ï¾ç´ßÁý¾ç³ä¼Ò½º|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910181|10117021|»ï¾ç´ßÁý¾ç³ä¼Ò½º|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910182|10117021|»ï¾ç´ßÁý¾ç³ä¼Ò½º|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910183|10117021|»ï¾ç´ßÁý¾ç³ä¼Ò½º|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910184|10117021|»ï¾ç´ßÁý¾ç³ä¼Ò½º|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910188|10117021|»ï¾ç´ßÁý¾ç³ä¼Ò½º|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180129|20180129|I0101|201801290036|WMD0910185|00216951|¼ÛÇý°æ|2008101671822354|861121|2AAEjAFRdDJ/dcYvj8o8rt9ND|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290036|WMD0910186|00216951|¼ÛÇý°æ|2008101671822354|861121|2AAEjAFRdDJ/dcYvj8o8rt9ND|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290036|WMD0910187|00216951|¼ÛÇý°æ|2008101671822354|861121|2AAEjAFRdDJ/dcYvj8o8rt9ND|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290037|WMD0910185|10118013|ÀÌ½ÂÀç|I010100000052327|841019|1AAEDWTDOLpBIjgwhDxQL5zX2|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290037|WMD0910186|10118013|ÀÌ½ÂÀç|I010100000052327|841019|1AAEDWTDOLpBIjgwhDxQL5zX2|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290037|WMD0910187|10118013|ÀÌ½ÂÀç|I010100000052327|841019|1AAEDWTDOLpBIjgwhDxQL5zX2|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290084|WMD0910185|10118018|Àü¼ÒÈñ|2008091615664689|930710|2AAEkiHJC73gU9QjNUgoqzxmA|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290084|WMD0910186|10118018|Àü¼ÒÈñ|2008091615664689|930710|2AAEkiHJC73gU9QjNUgoqzxmA|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290084|WMD0910187|10118018|Àü¼ÒÈñ|2008091615664689|930710|2AAEkiHJC73gU9QjNUgoqzxmA|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290085|WMD0910185|10118017|À±Âù¹Ì|I130100000375849|920629|2AAEGhRZybQ/G5kOev+0BjH+j|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290085|WMD0910186|10118017|À±Âù¹Ì|I130100000375849|920629|2AAEGhRZybQ/G5kOev+0BjH+j|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290085|WMD0910187|10118017|À±Âù¹Ì|I130100000375849|920629|2AAEGhRZybQ/G5kOev+0BjH+j|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290103|WMD0910185|10061809|ÃÖ¼øÈñ|2008091015474332|770902|2AAEnACI5oeWP0AzCceZUz9vD|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290103|WMD0910186|10061809|ÃÖ¼øÈñ|2008091015474332|770902|2AAEnACI5oeWP0AzCceZUz9vD|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290103|WMD0910187|10061809|ÃÖ¼øÈñ|2008091015474332|770902|2AAEnACI5oeWP0AzCceZUz9vD|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290148|WMD0910185|00054686|Á¶À±¿Á|2008061231073932|820210|2AAEfShJJOIfa9WmOBMDzQ9Zr|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290148|WMD0910186|00054686|Á¶À±¿Á|2008061231073932|820210|2AAEfShJJOIfa9WmOBMDzQ9Zr|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290148|WMD0910187|00054686|Á¶À±¿Á|2008061231073932|820210|2AAEfShJJOIfa9WmOBMDzQ9Zr|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290149|WMD0910185|10017135|¼­¼÷Èñ|2008061030133445|710418|2AAHC/eTSqqBqvs5ls41xYygl|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290149|WMD0910186|10017135|¼­¼÷Èñ|2008061030133445|710418|2AAHC/eTSqqBqvs5ls41xYygl|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290149|WMD0910187|10017135|¼­¼÷Èñ|2008061030133445|710418|2AAHC/eTSqqBqvs5ls41xYygl|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290157|WMD0910185|00033390|¹ÚÇý¿µ|2008092918367282|890131|2AAHGfNo0OLHgnwGk9FGvVr4F|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290157|WMD0910186|00033390|¹ÚÇý¿µ|2008092918367282|890131|2AAHGfNo0OLHgnwGk9FGvVr4F|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290157|WMD0910187|00033390|¹ÚÇý¿µ|2008092918367282|890131|2AAHGfNo0OLHgnwGk9FGvVr4F|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290159|WMD0910185|10016141|±è¼öºó|2007122700062926|880525|1AAGYRWKhoXOnk4VPINHxI7/+|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290159|WMD0910186|10016141|±è¼öºó|2007122700062926|880525|1AAGYRWKhoXOnk4VPINHxI7/+|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290159|WMD0910187|10016141|±è¼öºó|2007122700062926|880525|1AAGYRWKhoXOnk4VPINHxI7/+|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290206|WMD0910185|00047737|ÀÌ¹Ì¿¬|2008092918394552|780408|2AAHpCg8s6nXHwiEHlK/KCEQE|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290206|WMD0910186|00047737|ÀÌ¹Ì¿¬|2008092918394552|780408|2AAHpCg8s6nXHwiEHlK/KCEQE|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290206|WMD0910187|00047737|ÀÌ¹Ì¿¬|2008092918394552|780408|2AAHpCg8s6nXHwiEHlK/KCEQE|2|S2|"
    
    SetRawData "Worklist Return : " & vbCrLf & send
    
    Get_WorkList = send
    
    Set oSOAP = Nothing

    DoEvents
    
    Exit Function

ErrHandle:
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        SetRawData "[Err]" & Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
        SetRawData "[Err]" & Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
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
    
'send = ""
'send = send & "MSH|^~\&|HL7|MMS|||20180129212430||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & vbCr
'send = send & "PID|||201801290036^¼ÛÇý°æ^861121^2^20180129^20180129^DefaultDomain^PI" & vbCr
'send = send & "PV1||E|I0101" & vbCr
'send = send & "OBR|1||||||20180129212430" & vbCr
'send = send & "OBX|1|ST|WMD0910185||||||||R" & vbCr
'send = send & "OBX|2|ST|WMD0910186||||||||R" & vbCr
'send = send & "OBX|3|ST|WMD0910187||||||||R" & vbCr

    SetRawData "send : " & vbCrLf & send
    
    Get_OrderList = send
    
    'MsgBox send
    
   ' MsgBox send
    
    'SetRawData "[Rcv]" & Get_OrderList
    Set oSOAP = Nothing

    DoEvents
    
    Exit Function

ErrHandle:

    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        SetRawData "[Err]" & Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
        SetRawData "[Err]" & Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
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
    
'    SendResult = CInt(send)
    
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

