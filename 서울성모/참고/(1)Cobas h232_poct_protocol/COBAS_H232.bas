Attribute VB_Name = "COBAS_H232"
Option Explicit

Public gXML_Header As String
Public gSend_ControlID As Long
Public gSend_ControlID2 As Long
Public gRece_ControlID As String

Type XML_Data
    Rece_ControlID As String
    DataType As String
    DateTime As String
    Result As String
    EquipCode As String
    Barcode As String
    MethodCode As String
    StatusCode As String
    InterpretationCode As String
End Type

Type XML_Data1
    Rece_ControlID1 As String
    DataType1 As String
    DateTime1 As String
    Result1 As String
    EquipCode1 As String
    Barcode1 As String
    MethodCode1 As String
    StatusCode1 As String
    InterpretationCode1 As String
End Type

Public gXML As XML_Data
Public gXML1 As XML_Data1


Public Sub Xml_Init()
    gXML.Rece_ControlID = ""
    gXML.DataType = ""
    gXML.DateTime = ""
    gXML.Result = ""
    gXML.EquipCode = ""
    gXML.Barcode = ""
    gXML.MethodCode = ""
    gXML.StatusCode = ""
    gXML.InterpretationCode = ""
End Sub

Public Sub Xml_Init1()
    gXML1.Rece_ControlID1 = ""
    gXML1.DataType1 = ""
    gXML1.DateTime1 = ""
    gXML1.Result1 = ""
    gXML1.EquipCode1 = ""
    gXML1.Barcode1 = ""
    gXML1.MethodCode1 = ""
    gXML1.StatusCode1 = ""
    gXML1.InterpretationCode1 = ""
End Sub

Public Function WinSock_ACK(asControlID As String, asEquipNum As String) As String
    Dim strACK As String
    Dim strDateTime As String
    Dim strControl_ID As String
    
    WinSock_ACK = ""
    
    strControl_ID = CStr(Send_ControlID(asEquipNum))
    
    strACK = gXML_Header & vbLf
    strACK = strACK & "<ACK.R01>" & vbLf
    strACK = strACK & "<HDR>" & vbLf
    strACK = strACK & "<HDR.control_id V= """ & strControl_ID & """/>" & vbLf
    strACK = strACK & "<HDR.version_id V= ""POCT1""/>" & vbLf
    strACK = strACK & "<HDR.creation_dttm V= """ & WinSock_DateTime & """/>" & vbLf
    strACK = strACK & "</HDR>" & vbLf
    strACK = strACK & "<ACK>" & vbLf
    strACK = strACK & "<ACK.type_cd V=""AA""/>" & vbLf
    strACK = strACK & "<ACK.ack_control_id V=""" & asControlID & """/>" & vbLf
    strACK = strACK & "</ACK>" & vbLf
    strACK = strACK & "</ACK.R01>"
    WinSock_ACK = strACK
    
End Function

Public Function WinSock_END(asControlID As String, asEquipNum As String) As String
    Dim strACK As String
    Dim strDateTime As String
    Dim strControl_ID As String
    
    WinSock_END = ""
    
    strControl_ID = CStr(Send_ControlID(asEquipNum))
    
    strACK = gXML_Header & vbLf
    strACK = strACK & "<END.R01>" & vbLf
    strACK = strACK & "<HDR>" & vbLf
    strACK = strACK & "<HDR.control_id V= """ & strControl_ID & """/>" & vbLf
    strACK = strACK & "<HDR.version_id V= ""POCT1""/>" & vbLf
    strACK = strACK & "<HDR.creation_dttm V= """ & WinSock_DateTime & """/>" & vbLf
    strACK = strACK & "</HDR>" & vbLf
    strACK = strACK & "<TRM>" & vbLf
    strACK = strACK & "<TRM.reason_cd V=""NRM""/>" & vbLf
    strACK = strACK & "</TRM>" & vbLf
    strACK = strACK & "</END.R01>"
    WinSock_END = strACK
    
End Function

Public Function WinSock_REQ(asEquipNum As String) As String
    Dim strACK As String
    Dim strDateTime As String
    Dim strControl_ID As String
    
    WinSock_REQ = ""
    
    strControl_ID = CStr(Send_ControlID(asEquipNum))
    
    strACK = gXML_Header & vbLf
    strACK = strACK & "<REQ.R01>" & vbLf
    strACK = strACK & "<HDR>" & vbLf
    strACK = strACK & "<HDR.control_id V= """ & strControl_ID & """/>" & vbLf
    strACK = strACK & "<HDR.version_id V= ""POCT1""/>" & vbLf
    strACK = strACK & "<HDR.creation_dttm V= """ & WinSock_DateTime & """/>" & vbLf
    strACK = strACK & "</HDR>" & vbLf
    strACK = strACK & "<REQ>" & vbLf
    strACK = strACK & "<REQ.request_cd V=""ROBS""/>" & vbLf
    strACK = strACK & "</REQ>" & vbLf
    strACK = strACK & "</REQ.R01>"
    WinSock_REQ = strACK
End Function

Public Function WinSock_ACK1(asControlID As String, asEquipNum As String) As String
    Dim strACK As String
    Dim strDateTime As String
    Dim strControl_ID As String
    
    WinSock_ACK1 = ""
    
    strControl_ID = CStr(Send_ControlID1(asEquipNum))
    
    strACK = gXML_Header & vbLf
    strACK = strACK & "<ACK.R01>" & vbLf
    strACK = strACK & "<HDR>" & vbLf
    strACK = strACK & "<HDR.control_id V= """ & strControl_ID & """/>" & vbLf
    strACK = strACK & "<HDR.version_id V= ""POCT1""/>" & vbLf
    strACK = strACK & "<HDR.creation_dttm V= """ & WinSock_DateTime & """/>" & vbLf
    strACK = strACK & "</HDR>" & vbLf
    strACK = strACK & "<ACK>" & vbLf
    strACK = strACK & "<ACK.type_cd V=""AA""/>" & vbLf
    strACK = strACK & "<ACK.ack_control_id V=""" & asControlID & """/>" & vbLf
    strACK = strACK & "</ACK>" & vbLf
    strACK = strACK & "</ACK.R01>"
    WinSock_ACK1 = strACK
    
End Function

Public Function WinSock_END1(asControlID As String, asEquipNum As String) As String
    Dim strACK As String
    Dim strDateTime As String
    Dim strControl_ID As String
    
    WinSock_END1 = ""
    
    strControl_ID = CStr(Send_ControlID1(asEquipNum))
    
    strACK = gXML_Header & vbLf
    strACK = strACK & "<END.R01>" & vbLf
    strACK = strACK & "<HDR>" & vbLf
    strACK = strACK & "<HDR.control_id V= """ & strControl_ID & """/>" & vbLf
    strACK = strACK & "<HDR.version_id V= ""POCT1""/>" & vbLf
    strACK = strACK & "<HDR.creation_dttm V= """ & WinSock_DateTime & """/>" & vbLf
    strACK = strACK & "</HDR>" & vbLf
    strACK = strACK & "<TRM>" & vbLf
    strACK = strACK & "<TRM.reason_cd V=""NRM""/>" & vbLf
    strACK = strACK & "</TRM>" & vbLf
    strACK = strACK & "</END.R01>"
    WinSock_END1 = strACK
    
End Function

Public Function WinSock_REQ1(asEquipNum As String) As String
    Dim strACK As String
    Dim strDateTime As String
    Dim strControl_ID As String
    
    WinSock_REQ1 = ""
    
    strControl_ID = CStr(Send_ControlID1(asEquipNum))
    
    strACK = gXML_Header & vbLf
    strACK = strACK & "<REQ.R01>" & vbLf
    strACK = strACK & "<HDR>" & vbLf
    strACK = strACK & "<HDR.control_id V= """ & strControl_ID & """/>" & vbLf
    strACK = strACK & "<HDR.version_id V= ""POCT1""/>" & vbLf
    strACK = strACK & "<HDR.creation_dttm V= """ & WinSock_DateTime & """/>" & vbLf
    strACK = strACK & "</HDR>" & vbLf
    strACK = strACK & "<REQ>" & vbLf
    strACK = strACK & "<REQ.request_cd V=""ROBS""/>" & vbLf
    strACK = strACK & "</REQ>" & vbLf
    strACK = strACK & "</REQ.R01>"
    WinSock_REQ1 = strACK
End Function


Public Function WinSock_DateTime() As String
'2008-10-06T00:18:36+00:00
    WinSock_DateTime = Format(Date, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ss") & "+00:00"
    
End Function

'XML parsing==============================================================================================================
Public Function XML_Parsing(ByVal asData As String) As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    XML_Parsing = ""
    sFileName = "Res"
    sParam = asData
    
    Xml_Log sParam, sFileName
    
    If Dir(App.Path & "\" & sFileName & ".xml") <> "" Then
        DisplayNode App.Path & "\" & sFileName & ".xml"
    End If
    
    
End Function

'XML parsing==============================================================================================================
Public Function XML_Parsing1(ByVal asData As String) As String

    Dim sRetStr As String
    Dim sFileName As String
    Dim sParam As String
    
    XML_Parsing1 = ""
    sFileName = "Res1"
    sParam = asData
    
    Xml_Log sParam, sFileName
    
    If Dir(App.Path & "\" & sFileName & ".xml") <> "" Then
        DisplayNode1 App.Path & "\" & sFileName & ".xml"
    End If
    
    
End Function

Public Function Send_ControlID1(asEquipNum As String) As Long
    If asEquipNum = "1" Then
        gSend_ControlID = gSend_ControlID + 1
        If gSend_ControlID > 99999 Then
            gSend_ControlID = 10000
        End If
        
        Send_ControlID1 = gSend_ControlID
    Else
    
        gSend_ControlID2 = gSend_ControlID2 + 1
        If gSend_ControlID2 > 99999 Then
            gSend_ControlID2 = 10000
        End If
        
        Send_ControlID1 = gSend_ControlID2
    End If
End Function


Public Function Send_ControlID(asEquipNum As String) As Long
    If asEquipNum = "1" Then
        gSend_ControlID = gSend_ControlID + 1
        If gSend_ControlID > 99999 Then
            gSend_ControlID = 10000
        End If
        
        Send_ControlID = gSend_ControlID
    Else
    
        gSend_ControlID2 = gSend_ControlID2 + 1
        If gSend_ControlID2 > 99999 Then
            gSend_ControlID2 = 10000
        End If
        
        Send_ControlID = gSend_ControlID2
    End If
End Function

Public Sub DisplayNode(asPath As String)

    Dim xmlDoc As New MSXML2.DOMDocument30
    Dim nodeBook As IXMLDOMElement
    Dim nodeId As IXMLDOMAttribute
    Dim xNode As MSXML2.IXMLDOMNode
    Dim namedNodeMap As IXMLDOMNamedNodeMap
    Dim Child_Node As MSXML2.IXMLDOMNodeList
    Dim MsgType As String
    
    On Error GoTo ErrXML:
    
    Xml_Init
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
'''        MsgBox ("You have error " & myErr.reason)
    Else
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                MsgType = Mid(xNode.nodeName, 1, 5)
                If MsgType = "HEL.R" Or MsgType = "DST.R" Or MsgType = "OBS.R" Or MsgType = "EOT.R" Or MsgType = "END.R" Then
                    Exit For
                End If
                
            End If
        Next
        Set Child_Node = Nothing
        
        If MsgType = "HEL.R" Or MsgType = "DST.R" Or MsgType = "EOT.R" Or MsgType = "END.R" Then
            Set nodeBook = xmlDoc.selectSingleNode("//HDR.control_id")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML.Rece_ControlID = nodeId.Value
                gXML.DataType = MsgType
            End If
            
            Set nodeBook = Nothing
            Set nodeId = Nothing
            
        ElseIf MsgType = "OBS.R" Then

            gXML.DataType = MsgType
            
            'ControlID
            
            Set nodeBook = xmlDoc.selectSingleNode("//HDR.control_id")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML.Rece_ControlID = nodeId.Value
            End If
            
            Set nodeBook = Nothing
            Set nodeId = Nothing
            
            '검사일시
            Set nodeBook = xmlDoc.selectSingleNode("//SVC.observation_dttm")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML.DateTime = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            
            '바코드번호
            Set nodeBook = xmlDoc.selectSingleNode("//PT/PT.patient_id")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML.Barcode = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            '장비코드
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.observation_id")
            If TypeName(nodeBook.Attributes.getNamedItem("DN")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("DN")
                gXML.EquipCode = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            'Result
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.value")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML.Result = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            'method
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.method_cd ")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML.MethodCode = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            'state
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.status_cd")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML.StatusCode = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            'interpretation_cd
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.interpretation_cd")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML.InterpretationCode = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
        End If
        
    End If
ErrXML:
    Exit Sub
    
End Sub


Public Sub DisplayNode1(asPath As String)

    Dim xmlDoc As New MSXML2.DOMDocument30
    Dim nodeBook As IXMLDOMElement
    Dim nodeId As IXMLDOMAttribute
    Dim xNode As MSXML2.IXMLDOMNode
    Dim namedNodeMap As IXMLDOMNamedNodeMap
    Dim Child_Node As MSXML2.IXMLDOMNodeList
    Dim MsgType As String
    
    On Error GoTo ErrXML:
    
    Xml_Init1
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
'''        MsgBox ("You have error " & myErr.reason)
    Else
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                MsgType = Mid(xNode.nodeName, 1, 5)
                If MsgType = "HEL.R" Or MsgType = "DST.R" Or MsgType = "OBS.R" Or MsgType = "EOT.R" Or MsgType = "END.R" Then
                    Exit For
                End If
                
            End If
        Next
        Set Child_Node = Nothing
        
        If MsgType = "HEL.R" Or MsgType = "DST.R" Or MsgType = "EOT.R" Or MsgType = "END.R" Then
            Set nodeBook = xmlDoc.selectSingleNode("//HDR.control_id")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML1.Rece_ControlID1 = nodeId.Value
                gXML1.DataType1 = MsgType
            End If
            
            Set nodeBook = Nothing
            Set nodeId = Nothing
            
        ElseIf MsgType = "OBS.R" Then

            gXML1.DataType1 = MsgType
            
            'ControlID
            
            Set nodeBook = xmlDoc.selectSingleNode("//HDR.control_id")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML1.Rece_ControlID1 = nodeId.Value
            End If
            
            Set nodeBook = Nothing
            Set nodeId = Nothing
            
            '검사일시
            Set nodeBook = xmlDoc.selectSingleNode("//SVC.observation_dttm")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML1.DateTime1 = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            
            '바코드번호
            Set nodeBook = xmlDoc.selectSingleNode("//PT/PT.patient_id")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML1.Barcode1 = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            '장비코드
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.observation_id")
            If TypeName(nodeBook.Attributes.getNamedItem("DN")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("DN")
                gXML1.EquipCode1 = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            'Result
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.value")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML1.Result1 = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            'method
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.method_cd ")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML1.MethodCode1 = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            'state
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.status_cd")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML1.StatusCode1 = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            'interpretation_cd
            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.interpretation_cd")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                gXML1.InterpretationCode1 = nodeId.Value
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
        End If
        
    End If
ErrXML:
    Exit Sub
    
End Sub
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
    
    sFileName = argFileName
    If Dir(App.Path & "\" & sFileName & ".xml") <> "" Then
        Kill App.Path & "\" & sFileName & ".xml"
    End If
    
    Open App.Path & "\" & sFileName & ".xml" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

