Attribute VB_Name = "COBAS_H232"
Option Explicit

Public gXML_Header As String
Public gSend_ControlID As Long
Public gSend_ControlID2 As Long
Public gRece_ControlID As String

Type XML_Data
    ModelType As String
    DeviceName As String
    Rece_ControlID As String
    DataType As String
    DateTime As String
    Result(2) As String
    EquipCode(2) As String
'    Result_INR As String
'    Result_PER As String
'    Result_SEC As String
'    EquipCode_INR As String
'    EquipCode_PER As String
'    EquipCode_SEC As String
    Barcode As String
    MethodCode(2) As String
    StatusCode(2) As String
    InterpretationCode(2) As String
    normal_lohi_limit(2) As String
    '--QC
    role_cd    As String
    lot_number  As String
    level_cd    As String
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
Dim intIdx As Integer


Public Sub Xml_Init()
    gXML.Rece_ControlID = ""
    gXML.DataType = ""
    gXML.DateTime = ""
    gXML.Result(0) = ""
    gXML.Result(1) = ""
    gXML.Result(2) = ""
'    gXML.Result_INR = ""
'    gXML.Result_PER = ""
'    gXML.Result_SEC = ""
    gXML.EquipCode(0) = ""
    gXML.EquipCode(1) = ""
    gXML.EquipCode(2) = ""
'    gXML.EquipCode_INR = ""
'    gXML.EquipCode_PER = ""
'    gXML.EquipCode_SEC = ""
    gXML.Barcode = ""
    gXML.MethodCode(0) = ""
    gXML.MethodCode(1) = ""
    gXML.MethodCode(2) = ""
    gXML.StatusCode(0) = ""
    gXML.StatusCode(1) = ""
    gXML.StatusCode(2) = ""
    gXML.InterpretationCode(0) = ""
    gXML.InterpretationCode(1) = ""
    gXML.InterpretationCode(2) = ""
    gXML.normal_lohi_limit(0) = ""
    gXML.normal_lohi_limit(1) = ""
    gXML.normal_lohi_limit(2) = ""
    gXML.lot_number = ""
    gXML.level_cd = ""
    gXML.role_cd = ""
    intIdx = 0
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
    WinSock_DateTime = Format(Date, "yyyy-mm-dd") & "T" & Format(time, "hh:mm:ss") & "+00:00"
    
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

Public Sub display_online_parsing_Rece(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    Dim xNode As MSXML2.IXMLDOMNode
    

    For Each xNode In Nodes
        'Debug.Print xNode.nodeName
        'Debug.Print xNode.Attributes.Item(0).nodeValue
        Select Case Trim(xNode.nodeName)
            Case "OBS.observation_id"
                gXML.EquipCode(intIdx) = xNode.Attributes.Item(0).nodeValue
            Case "OBS.value"
                gXML.Result(intIdx) = xNode.Attributes.Item(0).nodeValue
            Case "OBS.method_cd"
                gXML.MethodCode(intIdx) = xNode.Attributes.Item(0).nodeValue
            Case "OBS.interpretation_cd"
                gXML.InterpretationCode(intIdx) = xNode.Attributes.Item(0).nodeValue
            Case "OBS.normal_lo-hi_limit"
                gXML.normal_lohi_limit(intIdx) = xNode.Attributes.Item(0).nodeValue
                intIdx = intIdx + 1
                If intIdx = 3 Then
                    intIdx = 0
                End If
            Case "RGT.lot_number"
                gXML.lot_number = xNode.Attributes.Item(0).nodeValue
            Case "CTC.level_cd"
                gXML.level_cd = xNode.Attributes.Item(0).nodeValue
            Case "SVC.role_cd"
                gXML.role_cd = xNode.Attributes.Item(0).nodeValue
                
        End Select

        If xNode.hasChildNodes Then
            display_online_parsing_Rece xNode.childNodes, Indent
        End If
        
    Next xNode
End Sub

Public Sub DisplayNode(asPath As String)

    Dim xmlDoc As New MSXML2.DOMDocument30
    Dim nodeBook As IXMLDOMElement
    Dim nodeId As IXMLDOMAttribute
    Dim xNode As MSXML2.IXMLDOMNode
    Dim namedNodeMap As IXMLDOMNamedNodeMap
    Dim Child_Node As MSXML2.IXMLDOMNodeList
    Dim MsgType As String
    Dim strBuffer As String
    Dim IntRow As Long
    Dim varBuffer As Variant
    Dim blnQc     As Boolean
    
    On Error GoTo ErrXML:
    
    Xml_Init
        
    Set xmlDoc = New MSXML2.DOMDocument30
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
'''        MsgBox ("You have error " & myErr.reason)
    Else
    
'        For Each xNode In xmlDoc.childNodes
'            Debug.Print Trim(xNode.parentNode.nodeName)
'        Next
        
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
                
                '모델타입/장비코드
                Set nodeBook = xmlDoc.selectSingleNode("//DEV.vendor_id")
                If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                    Set nodeId = nodeBook.Attributes.getNamedItem("V")
                    gXML.ModelType = mGetP(nodeId.Value, 2, "&")
                    gXML.DeviceName = mGetP(nodeId.Value, 3, "&")
                End If
                Set nodeBook = Nothing
                Set nodeId = Nothing
                
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
            
            'QC 구분
            Set nodeBook = xmlDoc.selectSingleNode("//SVC.role_cd")
            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                Set nodeId = nodeBook.Attributes.getNamedItem("V")
                If InStr(nodeId.Value, "QC") > 0 Then
                    blnQc = True
                Else
                    blnQc = False
                End If
            End If
            Set nodeBook = Nothing
            Set nodeId = Nothing
            
'' QC
''    <CTC>
''      <CTC.name V="ROCHE"/>
''      <CTC.lot_number V="89"/>
''      <CTC.expiration_date V="2013-02-28T23:59:59+00:00"/>
''      <CTC.level_cd V="1" SN="LEVEL"/>
            
'' 일반
''    <CTC>
''      <SVC.role_cd V="OBS"/>
''      <SVC.observation_dttm V="2012-03-12T16:41:37+00:00"/>
''      <SVC.sequence_nbr V=""/>
            
            If gXML.ModelType = "CoaguChekXSPlus" Then
                '일련번호
                If blnQc = True Then
                    Set nodeBook = xmlDoc.selectSingleNode("//CTC/CTC.lot_number")
                    If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                        Set nodeId = nodeBook.Attributes.getNamedItem("V")
                        gXML.Barcode = nodeId.Value
                    End If
                    Set nodeBook = Nothing
                    Set nodeId = Nothing
                Else
                    Set nodeBook = xmlDoc.selectSingleNode("//SVC/SVC.sequence_nbr")
                    If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                        Set nodeId = nodeBook.Attributes.getNamedItem("V")
                        gXML.Barcode = nodeId.Value
                    End If
                    Set nodeBook = Nothing
                    Set nodeId = Nothing
                End If
            Else
                '바코드번호
                If blnQc = True Then
                    Set nodeBook = xmlDoc.selectSingleNode("//CTC/CTC.lot_number")
                    If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                        Set nodeId = nodeBook.Attributes.getNamedItem("V")
                        gXML.Barcode = nodeId.Value
                    End If
                    Set nodeBook = Nothing
                    Set nodeId = Nothing
                Else
                    Set nodeBook = xmlDoc.selectSingleNode("//PT/PT.patient_id")
                    If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
                        Set nodeId = nodeBook.Attributes.getNamedItem("V")
                        gXML.Barcode = nodeId.Value
                    End If
                    Set nodeBook = Nothing
                    Set nodeId = Nothing
                End If
            End If

            '장비코드/결과
            display_online_parsing_Rece xmlDoc.childNodes, 0

'            '장비코드(INR)
'            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.observation_id")
'            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
'                Set nodeId = nodeBook.Attributes.getNamedItem("V")
'                gXML.EquipCode_INR = nodeId.Value
'            End If
'            Set nodeBook = Nothing
'            Set nodeId = Nothing
'
'            '장비코드(PER)
'            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.observation_id")
'            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
'                Set nodeId = nodeBook.Attributes.getNamedItem("V")
'                gXML.EquipCode_PER = nodeId.Value
'            End If
'            Set nodeBook = Nothing
'            Set nodeId = Nothing
'
'
'            'Result
'            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.value")
'            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
'                Set nodeId = nodeBook.Attributes.getNamedItem("V")
'                gXML.Result_INR = nodeId.Value
'            End If
'            Set nodeBook = Nothing
'            Set nodeId = Nothing
'            'method
'            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.method_cd ")
'            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
'                Set nodeId = nodeBook.Attributes.getNamedItem("V")
'                gXML.MethodCode = nodeId.Value
'            End If
'            Set nodeBook = Nothing
'            Set nodeId = Nothing
'            'state
'            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.status_cd")
'            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
'                Set nodeId = nodeBook.Attributes.getNamedItem("V")
'                gXML.StatusCode = nodeId.Value
'            End If
'            Set nodeBook = Nothing
'            Set nodeId = Nothing
'            'interpretation_cd
'            Set nodeBook = xmlDoc.selectSingleNode("//PT/OBS/OBS.interpretation_cd")
'            If TypeName(nodeBook.Attributes.getNamedItem("V")) <> "Nothing" Then
'                Set nodeId = nodeBook.Attributes.getNamedItem("V")
'                gXML.InterpretationCode = nodeId.Value
'            End If
'            Set nodeBook = Nothing
'            Set nodeId = Nothing
        End If
        
    End If
ErrXML:
    Exit Sub
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Private Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function

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
    Print #FilNum, time & " " & argSQL
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

