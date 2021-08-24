Attribute VB_Name = "modEhwaEMR"
Option Explicit

Public gOnline_Test As String
Public gServerPath As String
Public gIFUser As String
Public gIFName As String

Public giIndex  As Long

'�ϼ���
Public Const gXml_S01 = "PG_SRL.SLP91_S01"
Public Const gXml_S02 = "PG_SRL.SLP91_S02"
Public Const gXml_S03 = "PG_SRL.SLP91_S03"
Public Const gXml_S04 = "PG_SRL.SLP91_S04"
Public Const gXml_S07 = "PG_SRL.SLP91_S07"
Public Const gXml_S10 = "PG_SRL.SLP91_S10"
Public Const gXml_S18 = "PG_SRL.SLP91_S18"
Public Const gXml_S24 = "PG_SRL.SLP91_S24"
Public Const gXml_U07 = "PG_SRL.SLP91_U01"

'�̴��
'-- �α���
Public Const gXml_LOGIN = "PKG_MSE_LM_INTERFACE.PC_MSE_USER_SELECT"

'-- ���ڵ� ��ȸ
Public Const gXml_ORDER_SELECT = "PKG_MSE_LM_INTERFACE.PC_MSE_ORDER_SELECT"
'-- �������
Public Const gXml_RESULT_UPLOAD = "PKG_MSE_LM_INTERFACE.PC_MSE_INTERFACE_SAVE"

Public Const gXml_ACPT_SPNO = "PKG_MSE_LM_INTERFACE.PC_MSE_INS_ACPT_SPNO"

Public gOrderExam As String


Type Exam_Select
    TST_CD      As String
    TST_CNT     As Integer
End Type

Public gExam_Select()   As Exam_Select

Type PatInfo_Select
'    TST_CD      As String
'    TST_NM      As String
'    TST_FRCT_CD As String
'    ACPTNO_1    As String
'    PT_NO       As String
'    PT_NM       As String
'    Sex         As String
'    SPC_CD_1    As String
'    ORD_SITE    As String
'    TST_CLS     As String
'    RERUN       As String
'    TST_FRCT_CD1    As String
'    HSP_CLS     As String
'    ACPT_DTETM  As String
'    Age         As String
'    OK          As Integer

    SPCM_NO                 As String   '��ü��ȣ
    ACPT_DTM                As String   '��������
    EXM_ACPT_NO             As String   '������ȣ
    PT_NO                   As String   'ȯ�ڹ�ȣ
    PT_NM                   As String   'ȯ���̸�
    SEX_TP_CD               As String   '����
    PT_BRDY_DT              As String   '�������
    PT_HME_DEPT_CD          As String   'ȯ�������
    WD_DEPT_CD              As String   '����
    EXM_CD                  As String   '�˻��ڵ�
    TH1_SPCM_CD             As String   '��ü�ڵ�
    HR24_URN_EXM_TM         As String   '24�ð��Һ��˻�ð�
    HR24_URN_EXM_VLM_CNTE   As String   '24�ð��Һ��˻���ǳ���
    RPRN_EXM_CD             As String   '�ǳ��ڵ�
    EXM_PRGR_STS_CD         As String   '������°�
    OK                      As Integer
End Type

Public gPatInfo_Select  As PatInfo_Select

'======================== gnuh_emr ======================================

'Get_QCList ���ڵ��ȣ, ���� - QC ���� �ҷ�����
'Get_QCWorkList �˻�����, ����ȣ - QC WorkList �ҷ�����
'Online_QCResult "99910084349", "C061", "664887", "20091008151515", 5, " L63011  L63012  L63013  L6371   L6377   ", " 1.1 2.2 3.3 4.4 5.1 "
' - QC ��� ����

Type Order_Select
    SPC_NO      As String
    PT_NO       As String
    PT_NM       As String
    ACPT_DTE    As String
    ACPT_NO     As String
    TST_CD      As String
    WRK_UNT     As String
    TST_DTE     As String
    TST_STAT    As String
    WD_NO       As String
    SPC_NM      As String
    SPC_CD      As String
    OK          As Integer
End Type
Public gOrder_Select As Order_Select
Public gOrder_List() As Order_Select
Public gWork_Select() As Order_Select
'Public giIndex  As Long

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
    OK          As Integer
End Type
Public gPatient_Info As Patient_Info

'�߰� ���� start==========================================
Type QC_Info
    INST_DTM    As String
    LOT_NO      As String
    TST_CD      As String
    EQUIP_CD    As String
    CTRL_CD     As String
    LOT_NO1     As String
    BARCODE_CD  As String
    USE_YN      As String
    OK          As Integer
End Type

Public gOnline_Ret As String
'======================== gnuh_emr ======================================


Public Sub Clear_XML_Exam()
    giIndex = -1
    ReDim gExam_Select(0)
End Sub

Public Sub Clear_XML_PInfo()
'    gPatInfo_Select.ACPT_DTETM = ""
'    gPatInfo_Select.ACPTNO_1 = ""
'    gPatInfo_Select.Age = ""
'    gPatInfo_Select.HSP_CLS = ""
'    gPatInfo_Select.OK = -1
'    gPatInfo_Select.ORD_SITE = ""
'    gPatInfo_Select.PT_NM = ""
'    gPatInfo_Select.PT_NO = ""
'    gPatInfo_Select.RERUN = ""
'    gPatInfo_Select.Sex = ""
'    gPatInfo_Select.SPC_CD_1 = ""
'    gPatInfo_Select.TST_CD = ""
'    gPatInfo_Select.TST_CLS = ""
'    gPatInfo_Select.TST_FRCT_CD = ""
'    gPatInfo_Select.TST_FRCT_CD1 = ""
'    gPatInfo_Select.TST_NM = ""
    
    gPatInfo_Select.SPCM_NO = ""                       '��ü��ȣ
    gPatInfo_Select.ACPT_DTM = ""                      '��������
    gPatInfo_Select.EXM_ACPT_NO = ""                   '������ȣ
    gPatInfo_Select.PT_NO = ""                         'ȯ�ڹ�ȣ
    gPatInfo_Select.PT_NM = ""                         'ȯ���̸�
    gPatInfo_Select.SEX_TP_CD = ""                     '    ����
    gPatInfo_Select.PT_BRDY_DT = ""                    '�������
    gPatInfo_Select.PT_HME_DEPT_CD = ""                '�����
    gPatInfo_Select.WD_DEPT_CD = ""                    '����
    gPatInfo_Select.EXM_CD = ""                        '�˻��ڵ�
    gPatInfo_Select.OK = 0
    gPatInfo_Select.TH1_SPCM_CD = ""                   '��ü�ڵ�
    gPatInfo_Select.HR24_URN_EXM_TM = ""               '24Hour �ҹ��˻�ð�
    gPatInfo_Select.HR24_URN_EXM_VLM_CNTE = ""         '24Hour �ҹ��˻���ǳ���
    gPatInfo_Select.RPRN_EXM_CD = ""                   '�ǳ��ڵ�
    gPatInfo_Select.EXM_PRGR_STS_CD = ""               '������°�
    
End Sub

Public Function Online_TLA(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String

    Dim xDoc        As MSXML2.DOMDocument
    Dim xPE         As MSXML2.IXMLDOMParseError
    Dim sRetStr     As String
    Dim sFileName   As String
    Dim sParam      As String
    Dim strErrText  As String
    
    Online_TLA = ""
    sFileName = "Res"
    
    sParam = TLA_Param(asProc, asDate1, asDate2, asBarcode)
    
    sRetStr = Online_XML_Qry(asProc, sParam)
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    Set xDoc = New MSXML2.DOMDocument
    If xDoc.Load(App.PATH & "\XML\" & sFileName & ".xml") Then
        ' Data Load, Start Parsing
        Select Case asProc
        Case gXml_S04
'            Clear_XML_TLA
            display_online_parsing_TLAInfo xDoc.childNodes, 0
        Case gXml_S24
'            Clear_XML_TLA
            gIFName = ""
            display_online_parsing_User xDoc.childNodes, 0
        End Select
'        Display_Online_Parsing_Test xDoc.childNodes, 0
    Else
        ' ������ �ε����� ���߽��ϴ�.
       ' ParseError ��ü�� �����ɴϴ�
        Set xPE = xDoc.parseError
        With xPE
        
            strErrText = "Your XML Document failed to load" & _
                         "due the following error." & vbCrLf & _
                         "Error #: " & .errorCode & ": " & xPE.reason & _
                         "Line #: " & .Line & vbCrLf & _
                         "Line Position: " & .linepos & vbCrLf & _
                         "Position In File: " & .filepos & vbCrLf & _
                         "Source Text: " & .srcText & vbCrLf & _
                         "Document URL: " & .URL
        End With

'        SaveXML_Data strErrText
    End If
    
    Xml_Log gOnline_Test, "TLA"
    
    
    Set xPE = Nothing
    Set xDoc = Nothing
    
    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_TLA = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If
    
End Function

Public Function Online_XML(ByVal asProc As String, ByVal asSpcno As String, Optional ByVal asDIV As String, Optional ByVal PID As String, Optional ByVal PWD As String) As String
    Dim xmlDoc        As MSXML2.DOMDocument30
    Dim xPE         As MSXML2.IXMLDOMParseError
    Dim strErrText  As String
    Dim sRetStr     As String
    Dim sFileName   As String
    Dim sParam      As String
    Dim Nodes       As Object
    
    Dim nodeBook As IXMLDOMElement
    Dim nodeId As IXMLDOMAttribute
    Dim xNode As MSXML2.IXMLDOMNode
    Dim namedNodeMap As IXMLDOMNamedNodeMap
    Dim Child_Node As MSXML2.IXMLDOMNodeList
'    Dim MsgType As String
'    Dim strBuffer As String
'    Dim intRow As Long
'    Dim varBuffer As Variant
'    Dim blnQc     As Boolean
'    Dim i, J, k, m As Integer
'    Dim ii, jj, kk  As Integer
'    Dim strOData    As String
'    Dim strRData    As String
'
    Dim i As Integer
    Dim j As Integer
    
    
    Online_XML = ""
    sFileName = "Res"
    
    '�Ķ���͸� �����
    sParam = Select_Param(asProc, asSpcno, PID, PWD)
    
    '�Ķ���͸� ������ �����׿� ���ϰ��� �޾ƿ´�.
    sRetStr = Online_XML_Qry(asDIV, sParam)
    
    'sRetStr = "<?xml version='1.0' encoding='UTF-8'?>" & vbCrLf & sRetStr
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFileName
    
    
    Dim strData As String
    
    Erase strRecvData
    intBufCnt = 1000
    ReDim Preserve strRecvData(1)
    j = 1
    For i = 1 To Len(sRetStr)
        strData = Mid(sRetStr, i, 1)
        Select Case strData
            'Case "<"
            '    strRecvData(j) = strRecvData(j) & strData
            'Case "/"
            Case vbCr
            Case vbLf
                'strRecvData(j) = strRecvData(j) & strData
                j = j + 1
                ReDim Preserve strRecvData(j)
            Case Else
                
                strRecvData(j) = strRecvData(j) & strData
        End Select
    Next
    Set xmlDoc = New MSXML2.DOMDocument30
    
    
    'If xmlDoc.loadXML(App.PATH & "\XML\" & sFileName & ".xml") Then
    'If xmlDoc.loadXML(sRetStr) Then
    If UBound(strRecvData) > 1 Then
        
    '    Set Nodes = xmlDoc.documentElement
        
    '    If xmlDoc.readyState = 4 And xmlDoc.parseError.errorCode = 0 Then
    '        Online_XML = Nodes.Text
    '    Else
        
            ' Data Load, Start Parsing
            Select Case asProc
            Case gXml_S03
                Clear_XML_PInfo
                display_online_parsing_PatInfo xmlDoc.childNodes, 0
            Case gXml_S07
                Clear_XML_Exam
                display_online_parsing_ExamCode xmlDoc.childNodes, 0
            
            Case gXml_LOGIN
                'Clear_XML_PInfo
                'display_online_parsing_Login xmlDoc.childNodes, 0
                
                Online_XML = mGetP(mGetP(strRecvData(4), 2, "<"), 2, ">")
            Case gXml_ORDER_SELECT
                Clear_XML_PInfo
                'display_online_parsing_PatInfo xmlDoc.childNodes, 0
                display_online_PatInfo xmlDoc.childNodes, 0
                Online_XML = strRecvData(3)
            
            End Select
    '        Display_Online_Parsing_Test xmlDoc.childNodes, 0
            
     '   End If
    
    Else
        ' ������ �ε����� ���߽��ϴ�.
       ' ParseError ��ü�� �����ɴϴ�
        Set xPE = xmlDoc.parseError
        With xPE
        
            strErrText = "Your XML Document failed to load" & _
                         "due the following error." & vbCrLf & _
                         "Error #: " & .errorCode & ": " & xPE.reason & _
                         "Line #: " & .Line & vbCrLf & _
                         "Line Position: " & .linepos & vbCrLf & _
                         "Position In File: " & .filepos & vbCrLf & _
                         "Source Text: " & .srcText & vbCrLf & _
                         "Document URL: " & .URL
            
            Debug.Print strErrText
        
        End With

        SaveXML_Data strErrText
    End If
    
    
    Set xPE = Nothing
    Set xmlDoc = Nothing
    
    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_XML = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If
    
End Function

Public Function Online_XML_Qry(ByVal asStrDiv As String, ByVal asParam As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String
    
    On Error GoTo ErrHandle
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    oSOAP.MSSoapInit gHOSP.APIURL & "?wsdl"
    strDiv = asStrDiv
    sParam = asParam
    
    SaveXML_Data "[Use Proc => " & strDiv & " ]" & sParam
    send = oSOAP.LMService(strDiv, sParam)
    SaveXML_Data "[Return Proc => " & strDiv & " ]" & send
    Online_XML_Qry = send
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

Private Function Select_Param(ByVal asProc As String, ByVal asSpcno As String, Optional ByVal PID As String, Optional ByVal PWD As String) As String
    Dim sProc As String
    Dim sParam As String
    
    Select_Param = ""
    sParam = ""
    sProc = asProc
    
    Select Case sProc
    Case gXml_S01, gXml_S02, gXml_S03, gXml_S07, gXml_S10
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asSpcno & "]]></P0>" & _
                 "<P1><![CDATA[]]></P1>" & _
                 "</Table>"
    
    Case gXml_ORDER_SELECT
        sParam = ""
        sParam = sParam & "<?xml version='1.0' encoding='UTF-8'?>"
        sParam = sParam & "<Table>"
        sParam = sParam & "<QID><![CDATA[" & sProc & "]]></QID>"
        sParam = sParam & "<QTYPE><![CDATA[Package]]></QTYPE>"
        sParam = sParam & "<USERID><![CDATA[RTE]]></USERID>"
        sParam = sParam & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
        sParam = sParam & "<P0><![CDATA[" & mResult.SITE & "]]></P0>"
        sParam = sParam & "<P1><![CDATA[" & gHOSP.MACHCD & "]]></P1>"
        sParam = sParam & "<P2><![CDATA[]]></P2>"
        sParam = sParam & "<P3><![CDATA[]]></P3>"
        sParam = sParam & "<P4><![CDATA[]]></P4>"
        sParam = sParam & "<P5><![CDATA[]]></P5>"
        sParam = sParam & "<P6><![CDATA[]]></P6>"
        sParam = sParam & "<P7><![CDATA[" & asSpcno & "]]></P7>"
        sParam = sParam & "<P8><![CDATA[]]></P8>"
        sParam = sParam & "<P9><![CDATA[]]></P9>"
        sParam = sParam & "<P10><![CDATA[]]></P10>"
        sParam = sParam & "<P11><![CDATA[]]></P11>"
        sParam = sParam & "<P12><![CDATA[]]></P12>"
        sParam = sParam & "<P13><![CDATA[]]></P13>"
        sParam = sParam & "<P14><![CDATA[]]></P14>"
        sParam = sParam & "<P15><![CDATA[]]></P15>"
        sParam = sParam & "<P16><![CDATA[]]></P16>"
        sParam = sParam & "<P17><![CDATA[]]></P17>"
        sParam = sParam & "<P18><![CDATA[]]></P18>"
        sParam = sParam & "</Table>"
        
'<?xml version='1.0' encoding='UTF-8'?>
'<Table>
'<QID>
'<![CDATA[PKG_MSE_LM_INTERFACE.PC_MSE_USER_SELECT]]>
'</QID>
'<QTYPE>
'<![CDATA[Package]]>
'</QTYPE>
'<USERID>
'<![CDATA[RTE]]>
'</USERID>
'<EXECTYPE>
'<![CDATA[FILL]]>
'</EXECTYPE>
'<P0>
'<![CDATA[02]]>
'</P0>
'<P1>
'<![CDATA[C0EMR]]>
'</P1>
'<P2>
'<![CDATA[11111]]>
'</P2>
'</Table>
        
    Case gXml_LOGIN
        sParam = ""
        sParam = sParam & "<?xml version='1.0' encoding='UTF-8'?>"
        sParam = sParam & "<Table>"
        sParam = sParam & "<QID><![CDATA[" & sProc & "]]></QID>"
        sParam = sParam & "<QTYPE><![CDATA[Package]]></QTYPE>"
        sParam = sParam & "<USERID><![CDATA[RTE]]></USERID>"
        sParam = sParam & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
        sParam = sParam & "<P0><![CDATA[" & gHOSP.SITE & "]]></P0>"
        sParam = sParam & "<P1><![CDATA[" & PID & "]]></P1>"
        sParam = sParam & "<P2><![CDATA[" & PWD & "]]></P2>"
        sParam = sParam & "</Table>"
        
    End Select
    
    Select_Param = sParam
    
End Function

Private Function TLA_Param(ByVal asProc As String, ByVal asDate1 As String, ByVal asDate2 As String, Optional asBarcode As String = "0") As String
    Dim sProc As String
    Dim sParam As String
    
    TLA_Param = ""
    sProc = asProc
    
    Select Case sProc
    Case gXml_S04, gXml_S24
        sParam = "<Table>" & _
                 "<QID><![CDATA[" & sProc & "]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[LIA]]></USERID>" & _
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
                 "<USERID><![CDATA[LIA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & asBarcode & "]]></P0>" & _
                 "<P1><![CDATA[" & asDate1 & "]]></P1>" & _
                 "<P2><![CDATA[" & asDate2 & "]]></P2>" & _
                 "</Table>"
    End Select
    
    TLA_Param = sParam
    
End Function


'XML File Parsing===========================================================================================================
Public Sub display_online_parsing_ExamCode(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
'            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "TST_CD"
                giIndex = giIndex + 1
                ReDim Preserve gExam_Select(giIndex)
                gExam_Select(giIndex).TST_CD = xNode.nodeValue
                gExam_Select(giIndex).TST_CNT = giIndex + 1
                If gOrderExam = "" Then
                    gOrderExam = "'" & gExam_Select(giIndex).TST_CD & "'"
                Else
                    gOrderExam = gOrderExam & ", '" & gExam_Select(giIndex).TST_CD & "'"
                End If
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_ExamCode xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub display_online_parsing_PatInfo(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2
    giIndex = 0
    
    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
''            Select Case xNode.parentNode.nodeName
''            Case "TST_CD"
''                gPatInfo_Select.TST_CD = xNode.nodeValue
''                gPatInfo_Select.OK = 1
''            Case "ACPT_DTETM"
''                gPatInfo_Select.ACPT_DTETM = xNode.nodeValue
''            Case "ACPTNO_1"
''                gPatInfo_Select.ACPTNO_1 = xNode.nodeValue
''            Case "AGE"
''                gPatInfo_Select.Age = xNode.nodeValue
'''                Exit Sub
''            Case "HSP_CLS"
''                gPatInfo_Select.HSP_CLS = xNode.nodeValue
''            Case "ORD_SITE"
''                gPatInfo_Select.ORD_SITE = xNode.nodeValue
''            Case "PT_NM"
''                gPatInfo_Select.PT_NM = xNode.nodeValue
''            Case "PT_NO"
''                gPatInfo_Select.PT_NO = xNode.nodeValue
''            Case "RERUN"
''                gPatInfo_Select.RERUN = xNode.nodeValue
''            Case "SEX"
''                gPatInfo_Select.Sex = xNode.nodeValue
''            Case "SPC_CD_1"
''                gPatInfo_Select.SPC_CD_1 = xNode.nodeValue
''            Case "TST_CLS"
''                gPatInfo_Select.TST_CLS = xNode.nodeValue
''            Case "TST_FRCT_CD"
''                gPatInfo_Select.TST_FRCT_CD = xNode.nodeValue
''            Case "TST_FRCT_CD1"
''                gPatInfo_Select.TST_FRCT_CD1 = xNode.nodeValue
''            Case "TST_NM"
''                gPatInfo_Select.TST_NM = xNode.nodeValue
''            End Select
              
'              {
'              <?xml version="1.0" encoding="ISO-8859-1"?>
'              <string xmlns="http://tempuri.org/"><NewDataSet>
'              <Table>
'              <SPCM_NO>��ü��ȣ</SPCM_NO>
'              <ACPT_DTM>��������</ACPT_DTM>
'              <EXM_ACPT_NO>������ȣ</EXM_ACPT_NO>
'              <PT_NO>ȯ�ڹ�ȣ</PT_NO>
'              <PT_NM>ȯ���̸�</<PT_NO>
'              <SEX_TP_CD>����</SEX_TP_CD>
'              <PT_BRDY_DT>�������</PT_BRDY_DT>
'              <PT_HME_DEPT_CD>ȯ�������</PT_HME_DEPT_CD>
'              <WD_DEPT_CD>����</WD_DEPT_CD>
'              <EXM_CD>�˻��ڵ�</EXM_CD>
'              <TH1_SPCM_CD>��ü�ڵ�</TH1_SPCM_CD>
'              <HR24_URN_EXM_TM>24�ð��Һ��˻�ð�</HR24_URN_EXM_TM>
'              <HR24_URN_EXM_VLM_CNTE>24�ð��Һ��˻���ǳ���</HR24_URN_EXM_VLM_CNTE>
'              <RPRN_EXM_CD>�ǳ��ڵ�</RPRN_EXM_CD>
'              <EXM_PRGR_STS_CD>������°�</EXM_PRGR_STS_CD>
'              </Table>
'              </NewDataSet></string>
'              }
            
            Select Case xNode.parentNode.nodeName
                Case "SPCM_NO":                 gPatInfo_Select.SPCM_NO = xNode.nodeValue                       '��ü��ȣ
                Case "ACPT_DTM":                gPatInfo_Select.ACPT_DTM = xNode.nodeValue                      '��������
                Case "EXM_ACPT_NO":             gPatInfo_Select.EXM_ACPT_NO = xNode.nodeValue                   '������ȣ
                Case "PT_NO":                   gPatInfo_Select.PT_NO = xNode.nodeValue                         'ȯ�ڹ�ȣ
                Case "PT_NM":                   gPatInfo_Select.PT_NM = xNode.nodeValue                         'ȯ���̸�
                Case "SEX_TP_CD":               gPatInfo_Select.SEX_TP_CD = xNode.nodeValue                     '    ����
                Case "PT_BRDY_DT":              gPatInfo_Select.PT_BRDY_DT = xNode.nodeValue                    '�������
                Case "PT_HME_DEPT_CD":          gPatInfo_Select.PT_HME_DEPT_CD = xNode.nodeValue                '�����
                Case "WD_DEPT_CD":              gPatInfo_Select.WD_DEPT_CD = xNode.nodeValue                    '����
                Case "TH1_SPCM_CD":             gPatInfo_Select.TH1_SPCM_CD = xNode.nodeValue                   '��ü�ڵ�
                Case "HR24_URN_EXM_TM":         gPatInfo_Select.HR24_URN_EXM_TM = xNode.nodeValue               '24Hour �ҹ��˻�ð�
                Case "HR24_URN_EXM_VLM_CNTE":   gPatInfo_Select.HR24_URN_EXM_VLM_CNTE = xNode.nodeValue         '24Hour �ҹ��˻���ǳ���
                Case "RPRN_EXM_CD":             gPatInfo_Select.RPRN_EXM_CD = xNode.nodeValue                   '�ǳ��ڵ�
                Case "EXM_PRGR_STS_CD":         gPatInfo_Select.EXM_PRGR_STS_CD = xNode.nodeValue               '������°�
                Case "EXM_CD":                  gPatInfo_Select.EXM_CD = xNode.nodeValue                        '�˻��ڵ�
                    gPatInfo_Select.OK = 1
                    giIndex = giIndex + 1
                    ReDim Preserve gExam_Select(giIndex)
                    
                    gExam_Select(giIndex).TST_CD = xNode.nodeValue
                    gExam_Select(giIndex).TST_CNT = giIndex + 1
                    
                    If gPatOrdCd = "" Then
                        'gOrderExam = "'" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & "'" & gExam_Select(giIndex).TST_CD & "',"
                    Else
                        'gOrderExam = gOrderExam & ", '" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & ", '" & gExam_Select(giIndex).TST_CD & "'"
                    End If

                    gPatTest(giIndex) = gExam_Select(giIndex).TST_CD

            End Select
        
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_PatInfo xNode.childNodes, Indent
        End If
    Next xNode
    
    
End Sub

Public Sub display_online_PatInfo(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    Dim i               As Integer
    Dim strTemp         As String
    Dim strAtbName      As String
    Dim strAtbValue     As String
    
    Dim xNode As MSXML2.IXMLDOMNode
'    Indent = Indent + 2
    giIndex = 0
    
'    For Each xNode In Nodes
    
'        If xNode.nodeType = 4 Then
''            Select Case xNode.parentNode.nodeName
''            Case "TST_CD"
''                gPatInfo_Select.TST_CD = xNode.nodeValue
''                gPatInfo_Select.OK = 1
''            Case "ACPT_DTETM"
''                gPatInfo_Select.ACPT_DTETM = xNode.nodeValue
''            Case "ACPTNO_1"
''                gPatInfo_Select.ACPTNO_1 = xNode.nodeValue
''            Case "AGE"
''                gPatInfo_Select.Age = xNode.nodeValue
'''                Exit Sub
''            Case "HSP_CLS"
''                gPatInfo_Select.HSP_CLS = xNode.nodeValue
''            Case "ORD_SITE"
''                gPatInfo_Select.ORD_SITE = xNode.nodeValue
''            Case "PT_NM"
''                gPatInfo_Select.PT_NM = xNode.nodeValue
''            Case "PT_NO"
''                gPatInfo_Select.PT_NO = xNode.nodeValue
''            Case "RERUN"
''                gPatInfo_Select.RERUN = xNode.nodeValue
''            Case "SEX"
''                gPatInfo_Select.Sex = xNode.nodeValue
''            Case "SPC_CD_1"
''                gPatInfo_Select.SPC_CD_1 = xNode.nodeValue
''            Case "TST_CLS"
''                gPatInfo_Select.TST_CLS = xNode.nodeValue
''            Case "TST_FRCT_CD"
''                gPatInfo_Select.TST_FRCT_CD = xNode.nodeValue
''            Case "TST_FRCT_CD1"
''                gPatInfo_Select.TST_FRCT_CD1 = xNode.nodeValue
''            Case "TST_NM"
''                gPatInfo_Select.TST_NM = xNode.nodeValue
''            End Select
              
'              {
'              <?xml version="1.0" encoding="ISO-8859-1"?>
'              <string xmlns="http://tempuri.org/"><NewDataSet>
'              <Table>
'              <SPCM_NO>��ü��ȣ</SPCM_NO>
'              <ACPT_DTM>��������</ACPT_DTM>
'              <EXM_ACPT_NO>������ȣ</EXM_ACPT_NO>
'              <PT_NO>ȯ�ڹ�ȣ</PT_NO>
'              <PT_NM>ȯ���̸�</<PT_NO>
'              <SEX_TP_CD>����</SEX_TP_CD>
'              <PT_BRDY_DT>�������</PT_BRDY_DT>
'              <PT_HME_DEPT_CD>ȯ�������</PT_HME_DEPT_CD>
'              <WD_DEPT_CD>����</WD_DEPT_CD>
'              <EXM_CD>�˻��ڵ�</EXM_CD>
'              <TH1_SPCM_CD>��ü�ڵ�</TH1_SPCM_CD>
'              <HR24_URN_EXM_TM>24�ð��Һ��˻�ð�</HR24_URN_EXM_TM>
'              <HR24_URN_EXM_VLM_CNTE>24�ð��Һ��˻���ǳ���</HR24_URN_EXM_VLM_CNTE>
'              <RPRN_EXM_CD>�ǳ��ڵ�</RPRN_EXM_CD>
'              <EXM_PRGR_STS_CD>������°�</EXM_PRGR_STS_CD>
'              </Table>
'              </NewDataSet></string>
'              }
                    
        For i = 3 To UBound(strRecvData)
                        
            strTemp = mGetP(strRecvData(i), 2, "<")
            strAtbName = mGetP(strTemp, 1, ">")
            strAtbValue = mGetP(strTemp, 2, ">")
            
            Select Case strAtbName
                Case "SPCM_NO":                 gPatInfo_Select.SPCM_NO = strAtbValue                       '��ü��ȣ
                Case "ACPT_DTM":                gPatInfo_Select.ACPT_DTM = strAtbValue                      '��������
                Case "EXM_ACPT_NO":             gPatInfo_Select.EXM_ACPT_NO = strAtbValue                   '������ȣ
                Case "PT_NO":                   gPatInfo_Select.PT_NO = strAtbValue                         'ȯ�ڹ�ȣ
                Case "PT_NM":                   gPatInfo_Select.PT_NM = strAtbValue                         'ȯ���̸�
                Case "SEX_TP_CD":               gPatInfo_Select.SEX_TP_CD = strAtbValue                     '    ����
                Case "PT_BRDY_DT":              gPatInfo_Select.PT_BRDY_DT = strAtbValue                    '�������
                Case "PT_HME_DEPT_CD":          gPatInfo_Select.PT_HME_DEPT_CD = strAtbValue                '�����
                Case "WD_DEPT_CD":              gPatInfo_Select.WD_DEPT_CD = strAtbValue                    '����
                Case "TH1_SPCM_CD":             gPatInfo_Select.TH1_SPCM_CD = strAtbValue                   '��ü�ڵ�
                Case "HR24_URN_EXM_TM":         gPatInfo_Select.HR24_URN_EXM_TM = strAtbValue               '24Hour �ҹ��˻�ð�
                Case "HR24_URN_EXM_VLM_CNTE":   gPatInfo_Select.HR24_URN_EXM_VLM_CNTE = strAtbValue         '24Hour �ҹ��˻���ǳ���
                Case "RPRN_EXM_CD":             gPatInfo_Select.RPRN_EXM_CD = strAtbValue                   '�ǳ��ڵ�
                Case "EXM_PRGR_STS_CD":         gPatInfo_Select.EXM_PRGR_STS_CD = strAtbValue               '������°�
                Case "EXM_CD":                  gPatInfo_Select.EXM_CD = strAtbValue                        '�˻��ڵ�
                    gPatInfo_Select.OK = 1
                    giIndex = giIndex + 1
                    ReDim Preserve gExam_Select(giIndex)
                    ReDim Preserve gPatTest(giIndex)
                    
                    gExam_Select(giIndex).TST_CD = strAtbValue
                    gExam_Select(giIndex).TST_CNT = giIndex + 1
                    
                    If gPatOrdCd = "" Then
                        'gOrderExam = "'" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & "'" & gExam_Select(giIndex).TST_CD & "',"
                    Else
                        'gOrderExam = gOrderExam & ", '" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & ", '" & gExam_Select(giIndex).TST_CD & "'"
                    End If

                    gPatTest(giIndex) = gExam_Select(giIndex).TST_CD

            End Select
        Next
        'End If
        'If xNode.hasChildNodes Then
        '    display_online_parsing_PatInfo xNode.childNodes, Indent
        'End If
    'Next xNode
    
    

End Sub

Public Sub display_online_parsing_Login(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2
    giIndex = 0
    
    For Each xNode In Nodes
    
        If xNode.nodeType = 1 Then

            Select Case xNode.parentNode.nodeName
                Case "SPCM_NO":                 gPatInfo_Select.SPCM_NO = xNode.nodeValue                       '��ü��ȣ
                Case "ACPT_DTM":                gPatInfo_Select.ACPT_DTM = xNode.nodeValue                      '��������
                Case "EXM_ACPT_NO":             gPatInfo_Select.EXM_ACPT_NO = xNode.nodeValue                   '������ȣ
                Case "PT_NO":                   gPatInfo_Select.PT_NO = xNode.nodeValue                         'ȯ�ڹ�ȣ
                Case "PT_NM":                   gPatInfo_Select.PT_NM = xNode.nodeValue                         'ȯ���̸�
                Case "SEX_TP_CD":               gPatInfo_Select.SEX_TP_CD = xNode.nodeValue                     '    ����
                Case "PT_BRDY_DT":              gPatInfo_Select.PT_BRDY_DT = xNode.nodeValue                    '�������
                Case "PT_HME_DEPT_CD":          gPatInfo_Select.PT_HME_DEPT_CD = xNode.nodeValue                '�����
                Case "WD_DEPT_CD":              gPatInfo_Select.WD_DEPT_CD = xNode.nodeValue                    '����
                Case "TH1_SPCM_CD":             gPatInfo_Select.TH1_SPCM_CD = xNode.nodeValue                   '��ü�ڵ�
                Case "HR24_URN_EXM_TM":         gPatInfo_Select.HR24_URN_EXM_TM = xNode.nodeValue               '24Hour �ҹ��˻�ð�
                Case "HR24_URN_EXM_VLM_CNTE":   gPatInfo_Select.HR24_URN_EXM_VLM_CNTE = xNode.nodeValue         '24Hour �ҹ��˻���ǳ���
                Case "RPRN_EXM_CD":             gPatInfo_Select.RPRN_EXM_CD = xNode.nodeValue                   '�ǳ��ڵ�
                Case "EXM_PRGR_STS_CD":         gPatInfo_Select.EXM_PRGR_STS_CD = xNode.nodeValue               '������°�
                Case "EXM_CD":                  gPatInfo_Select.EXM_CD = xNode.nodeValue                        '�˻��ڵ�
                    gPatInfo_Select.OK = 1
                    giIndex = giIndex + 1
                    ReDim Preserve gExam_Select(giIndex)
                    
                    gExam_Select(giIndex).TST_CD = xNode.nodeValue
                    gExam_Select(giIndex).TST_CNT = giIndex + 1
                    
                    If gPatOrdCd = "" Then
                        'gOrderExam = "'" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & "'" & gExam_Select(giIndex).TST_CD & "',"
                    Else
                        'gOrderExam = gOrderExam & ", '" & gExam_Select(giIndex).TST_CD & "'"
                        gPatOrdCd = gPatOrdCd & ", '" & gExam_Select(giIndex).TST_CD & "'"
                    End If

                    gPatTest(giIndex) = gExam_Select(giIndex).TST_CD

            End Select
        
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_Login xNode.childNodes, Indent
        End If
    Next xNode
End Sub


Public Sub display_online_parsing_User(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
            Select Case xNode.parentNode.nodeName
            Case "NM"
                gIFName = xNode.nodeValue
            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_User xNode.childNodes, Indent
        End If
    Next xNode
End Sub


Public Sub display_online_parsing_TLAInfo(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab
'            Select Case xNode.parentNode.nodeName
'            Case "TST_CD"
'                gPatInfo_Select.TST_CD = xNode.nodeValue
'                gPatInfo_Select.OK = 1
'            Case "ACPT_DTETM"
'                gPatInfo_Select.ACPT_DTETM = xNode.nodeValue
'            Case "ACPTNO_1"
'                gPatInfo_Select.ACPTNO_1 = xNode.nodeValue
'            Case "AGE"
'                gPatInfo_Select.AGE = xNode.nodeValue
''                Exit Sub
'            Case "HSP_CLS"
'                gPatInfo_Select.HSP_CLS = xNode.nodeValue
'            Case "ORD_SITE"
'                gPatInfo_Select.ORD_SITE = xNode.nodeValue
'            Case "PT_NM"
'                gPatInfo_Select.PT_NM = xNode.nodeValue
'            Case "PT_NO"
'                gPatInfo_Select.PT_NO = xNode.nodeValue
'            Case "RERUN"
'                gPatInfo_Select.RERUN = xNode.nodeValue
'            Case "SEX"
'                gPatInfo_Select.SEX = xNode.nodeValue
'            Case "SPC_CD_1"
'                gPatInfo_Select.SPC_CD_1 = xNode.nodeValue
'            Case "TST_CLS"
'                gPatInfo_Select.TST_CLS = xNode.nodeValue
'            Case "TST_FRCT_CD"
'                gPatInfo_Select.TST_FRCT_CD = xNode.nodeValue
'            Case "TST_FRCT_CD1"
'                gPatInfo_Select.TST_FRCT_CD1 = xNode.nodeValue
'            Case "TST_NM"
'                gPatInfo_Select.TST_NM = xNode.nodeValue
'            End Select
        End If
        If xNode.hasChildNodes Then
            display_online_parsing_TLAInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub Display_Online_Parsing(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)
    
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
    
        If xNode.nodeType = 4 Then
            gOnline_Test = gOnline_Test & xNode.nodeValue & vbTab

        End If
        If xNode.hasChildNodes Then
            display_online_parsing_TLAInfo xNode.childNodes, Indent
        End If
    Next xNode
End Sub



'===========================================================================================================XML File Parsing

'Result Trans sub start ====================================================================================================
Public Function Online_Result(ByVal asParam As String) As String

    Dim sRetStr As String


    Online_Result = ""

    gOnline_Ret = ""

    sRetStr = Online_Result_Qry(asParam)

    'SaveXMLFile sRetStr
    Xml_Log sRetStr, "res"

    Dim xDoc As MSXML2.DOMDocument

    Set xDoc = New MSXML2.DOMDocument

    If xDoc.Load(App.PATH & "\Res\res.xml") Then
    'If xDoc.Load(sRetStr) Then
        ' ������ ���������� �ε�Ǿ����ϴ�.
        ' ���� ����ִ� �۾��� �����մϴ�.
        Display_Online_Parsing xDoc.childNodes, 0
    Else
        ' ������ �ε����� ���߽��ϴ�.
        Dim strErrText As String
        Dim xPE As MSXML2.IXMLDOMParseError
       ' ParseError ��ü�� �����ɴϴ�
        Set xPE = xDoc.parseError
        With xPE
            strErrText = "Your XML Document failed to load" & _
                         "due the following error." & vbCrLf & _
                         "Error #: " & .errorCode & ": " & xPE.reason & _
                         "Line #: " & .Line & vbCrLf & _
                         "Line Position: " & .linepos & vbCrLf & _
                         "Position In File: " & .filepos & vbCrLf & _
                         "Source Text: " & .srcText & vbCrLf & _
                         "Document URL: " & .URL
        End With

        SaveData strErrText
    End If

    Set xPE = Nothing

    Set xDoc = Nothing

    If InStr(1, gOnline_Ret, vbTab) > 0 Then
        Online_Result = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
    End If

End Function

Public Function Online_Result_Qry(ByVal asParam As String) As String
    Dim oSOAP As MSSOAPLib30.SoapClient30
    Dim strDiv As String
    Dim send As String
    Dim sParam As String

    On Error GoTo ErrHandle

    Set oSOAP = New MSSOAPLib30.SoapClient30

    oSOAP.ClientProperty("ServerHTTPRequest") = True
    
    oSOAP.MSSoapInit gHOSP.APIURL & "?wsdl"
    
    strDiv = gXml_RESULT_UPLOAD 'PKG_MSE_LM_INTERFACE.PC_MSE_INTERFACE_SAVE

    
    sParam = asParam
    
    SaveXML_Data "[Save Result]" & sParam
    'send = oSOAP.wsLISInterface(strDiv, sParam)
    send = oSOAP.LMService("SETQUERY", sParam)
    
    SaveXML_Data "[Save Result => Return]" & send
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

'==================================================================================================== Result Trans sub start

'������ ����=================================================================================================================
Public Sub SaveXML_Data(argSQL As String)
'argSQL�� ������ ���Ϸ� ����
    Dim FilNum

    FilNum = FreeFile

    If Dir(App.PATH & "\" & "XML", vbDirectory) <> "XML" Then
        MkDir (App.PATH & "\XML")
    End If

    Open App.PATH & "\XML" & "\" & Date & ".log" For Append As FilNum
    Print #FilNum, Time & " " & argSQL
    Close FilNum
End Sub

Public Sub Xml_Log(argSQL As String, argFileName As String)
'argSQL�� ������ ���Ϸ� ����
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.PATH & "\" & "XML", vbDirectory) <> "XML" Then
        MkDir (App.PATH & "\" & "XML")
    End If
    
    sFileName = argFileName
    If Dir(App.PATH & "\" & "XML" & "\" & sFileName & ".xml") <> "" Then
        Kill App.PATH & "\" & "XML" & "\" & sFileName & ".xml"
    End If
    
    Open App.PATH & "\" & "XML" & "\" & sFileName & ".xml" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

'=================================================================================================================������ ����



