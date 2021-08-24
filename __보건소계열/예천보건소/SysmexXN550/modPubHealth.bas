Attribute VB_Name = "modPubHealth"
Option Explicit

Public Function Get_WorkList() As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim send
    Dim sParam
    Dim sRet
    
    On Error GoTo ErrHandle
    
'    send = ""
'    send = send & "20160620|20160620|N1701|201606200014|WB1050|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WB1040|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WB2590|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WB2710|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC2210|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC2411|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC2420|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC2430|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC2443|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC3720|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC3721|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC3730|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC3750|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200014|WC3780|00028319|À¯¿µ¼÷|2008092218025794|660205|2AAEH5AzCa/Nbahrnjg+CfnvC|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WB2570|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WB2580|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WB2590|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WB2710|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC2210|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC2411|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC2420|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC2430|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC2443|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC3720|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC3721|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC3730|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC3750|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200021|WC3780|00005631|ÀüÁ¾¹é|2008092218051932|531118|1AAGa75oicK2kikfhzL+tox23|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WB2570|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WB2580|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WB2590|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WB2710|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC2210|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC2411|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC2420|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC2430|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC2443|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC3720|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC3721|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC3730|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC3750|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200022|WC3780|00013151|±è³²¼÷|2008092218041605|551020|2AAGxLEscejfrd7qrXvQaic+0|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200037|WB2570|00006816|±è¼ö¹Ì|2008092218041957|931008|2AAFjYNRtSdYq4dCntVI7iNqS|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200037|WB2580|00006816|±è¼ö¹Ì|2008092218041957|931008|2AAFjYNRtSdYq4dCntVI7iNqS|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WB1050|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WB1040|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WB2590|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WB2710|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC2210|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC2411|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC2420|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC2430|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC2443|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC3720|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC3721|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC3730|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC3750|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200038|WC3780|10011655|ÃÖÈ«ÀÓ|2008092218021887|381101|2AAGs1ep5V+rDpA9sIgztSfnn|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200040|WB2570|00024398|¹ÚºÐ¿µ|2008092218061300|620618|2AAGBLrlXs1PtzqJ10BRClwWi|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200040|WB2580|00024398|¹ÚºÐ¿µ|2008092218061300|620618|2AAGBLrlXs1PtzqJ10BRClwWi|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200041|WB2570|00002055|±èÇØ¿Á|2008092218054759|521202|2AAEQmNDzCSts6AstCRcLTpLs|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200041|WB2580|00002055|±èÇØ¿Á|2008092218054759|521202|2AAEQmNDzCSts6AstCRcLTpLs|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200042|WB2570|00029828|±è¹ÌÀÚ|2008092218035582|580415|2AAGcgaoS+ucw2CtevXcXylzE|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200042|WB2580|00029828|±è¹ÌÀÚ|2008092218035582|580415|2AAGcgaoS+ucw2CtevXcXylzE|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200043|WB2570|10021317|È²¼÷ÀÚ|2008092218053009|610205|2AAENuF2QA+rcCykSnbCZCDwZ|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200043|WB2580|10021317|È²¼÷ÀÚ|2008092218053009|610205|2AAENuF2QA+rcCykSnbCZCDwZ|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200044|WB2570|10021316|±èÁ¤Èñ|2008031824939602|560914|2AAG+NS9MkBfddOlzi2Sa9wMt|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200044|WB2580|10021316|±èÁ¤Èñ|2008031824939602|560914|2AAG+NS9MkBfddOlzi2Sa9wMt|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200045|WB2570|10021315|±Ç¿µ¼ø|2009122813438410|610802|2AAG2ANoKC0NuclGtPTNYIx5w|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200045|WB2580|10021315|±Ç¿µ¼ø|2009122813438410|610802|2AAG2ANoKC0NuclGtPTNYIx5w|2|C1|"
'    send = send & "20160620|20160620|N1701|201606200053|WB2570|10003173|ÃÖÁß½É|2008092218028101|910402|1AAEuCAZpW6l2c4N/D95Mq/dL|1|C1|"
'    send = send & "20160620|20160620|N1701|201606200053|WB2580|10003173|ÃÖÁß½É|2008092218028101|910402|1AAEuCAZpW6l2c4N/D95Mq/dL|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WB2570|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WB2580|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WB2590|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WB2710|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC2210|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC2411|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC2420|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC2430|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC2443|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC3720|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC3721|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC3730|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC3750|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"
'    send = send & "20160621|20160621|N1701|201606210003|WC3780|00027224|À±Ä¡»ç|2008092218050712|420218|1AAFWDv6EiBOAgwbEEjbKiY8E|1|C1|"

    
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
    
    If send = "" Then
        Set oSOAP = Nothing
        Get_WorkList = ""
        Exit Function
    End If
    
    send = makeUB64(send)
    
    SetRawData "Worklist Return : " & vbCrLf & send
    
    Get_WorkList = send
    
    Set oSOAP = Nothing

    DoEvents
    
    Exit Function

ErrHandle:
    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
        SetRawData "Get_WorkList:" & Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
        SetRawData "Get_WorkList:" & Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If
    
End Function

Public Function Get_OrderList(asSID As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim send As String
    Dim sParam As String
    
    Get_OrderList = ""
    
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
    'SetRawData "[sParam]" & sParam
    
    send = oSOAP.New_SelectOrder(sParam)
    'SetRawData "[send1]" & send
    If send = "" Then
        Set oSOAP = Nothing
        Get_OrderList = ""
        Exit Function
    End If
    
    send = makeUB64(send)
    
    If send = "" Then
        Set oSOAP = Nothing
        Get_OrderList = ""
        Exit Function
    End If
    
    'SetRawData "[send2]" & send
    
    SetRawData "New_SelectOrder Return : " & vbCrLf & send
    
    Get_OrderList = send
    
'    SetRawData "[Rcv]" & Get_OrderList
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
    
    Set oSOAP = Nothing
    Get_OrderList = ""
    
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
    
'    SetRawData "UpdateRst Ret : " & vbCrLf & send

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

