Module modWebData

    Public Function DBSysDate() As Date
        DBSysDate = Now
    End Function

    '검체번호로 검사코드 조회
    Public Function get_WEB_INTERFACE_S03(ByVal strSpcNo As String) As ADODB.Recordset
        Dim strDiv As String
        Dim strParam As String
        Dim objLIS As webLIS.LisInterface
        Dim strReturn As String

        strDiv = "PG_SRL.INTERFACE_S03"

        strParam = "<Table>" & _
                          "<QID><![CDATA[PG_SRL.INTERFACE_S03]]></QID>" & _
                          "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                          "<USERID><![CDATA[" & gstrConnectTyp & "]]></USERID>" & _
                          "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                          "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                          "<P0><![CDATA[" & strSpcNo & "]]></P0>" & _
                          "<P1><![CDATA[" & "" & "]]></P1>" & _
                   "</Table>"

        Try
            objLIS = New webLIS.LisInterface

            strReturn = objLIS.wsLISInterface(strDiv, strParam)

            get_WEB_INTERFACE_S03 = ConvertXmlToRecordSet(strReturn)

        Catch ex As Exception
            get_WEB_INTERFACE_S03 = Nothing

            ErrMsgProc(strDiv & vbNewLine & strParam)

        Finally
            If Not objLIS Is Nothing Then objLIS.Dispose()

            objLIS = Nothing
        End Try
    End Function

    '검체번호로 결과완료 되지 않은 검사코드 리스트 조회
    Public Function get_WEB_INTERFACE_S07(ByVal strSpcNo As String) As ADODB.Recordset
        Dim strDiv As String
        Dim strParam As String
        Dim objLIS As webLIS.LisInterface
        Dim strReturn As String

        strDiv = "PG_SRL.INTERFACE_S07"

        strParam = "<Table>" & _
                          "<QID><![CDATA[PG_SRL.INTERFACE_S07]]></QID>" & _
                          "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                          "<USERID><![CDATA[" & gstrConnectTyp & "]]></USERID>" & _
                          "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                          "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                          "<P0><![CDATA[" & strSpcNo & "]]></P0>" & _
                          "<P1><![CDATA[" & "" & "]]></P1>" & _
                   "</Table>"

        Try
            objLIS = New webLIS.LisInterface

            strReturn = objLIS.wsLISInterface(strDiv, strParam)

            'TST_CD
            get_WEB_INTERFACE_S07 = ConvertXmlToRecordSet(strReturn)

        Catch ex As Exception
            get_WEB_INTERFACE_S07 = Nothing

            ErrMsgProc(strDiv & vbNewLine & strParam)

        Finally
            If Not objLIS Is Nothing Then objLIS.Dispose()

            objLIS = Nothing
        End Try
    End Function

    '검체번호로 인터페이스 되지 않은 검사코드 리스트 조회
    Public Function get_WEB_INTERFACE_S08(ByVal strSpcNo As String) As ADODB.Recordset
        Dim strDiv As String
        Dim strParam As String
        Dim objLIS As webLIS.LisInterface
        Dim strReturn As String

        strDiv = "PG_SRL.INTERFACE_S08"

        strParam = "<Table>" & _
                          "<QID><![CDATA[PG_SRL.INTERFACE_S08]]></QID>" & _
                          "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                          "<USERID><![CDATA[" & gstrConnectTyp & "]]></USERID>" & _
                          "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                          "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                          "<P0><![CDATA[" & strSpcNo & "]]></P0>" & _
                          "<P1><![CDATA[" & "" & "]]></P1>" & _
                   "</Table>"

        Try
            objLIS = New webLIS.LisInterface

            strReturn = objLIS.wsLISInterface(strDiv, strParam)

            'TST_CD
            get_WEB_INTERFACE_S08 = ConvertXmlToRecordSet(strReturn)

        Catch ex As Exception
            get_WEB_INTERFACE_S08 = Nothing

            ErrMsgProc(strDiv & vbNewLine & strParam)

        Finally
            If Not objLIS Is Nothing Then objLIS.Dispose()

            objLIS = Nothing
        End Try
    End Function

    '검체번호로 환자정보 조회(환자번호, 이름, 성별, 나이, 진료과, 병동번호)
    Public Function get_WEB_INTERFACE_S06(ByVal strSpcNo As String) As ADODB.Recordset
        Dim strDiv As String
        Dim strParam As String
        Dim objLIS As webLIS.LisInterface
        Dim strReturn As String

        strDiv = "PG_SRL.INTERFACE_S06"

        strParam = "<Table>" & _
                          "<QID><![CDATA[PG_SRL.INTERFACE_S06]]></QID>" & _
                          "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                          "<USERID><![CDATA[" & gstrConnectTyp & "]]></USERID>" & _
                          "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                          "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                          "<P0><![CDATA[" & strSpcNo & "]]></P0>" & _
                          "<P1><![CDATA[" & "" & "]]></P1>" & _
                   "</Table>"

        Try
            objLIS = New webLIS.LisInterface

            strReturn = objLIS.wsLISInterface(strDiv, strParam)

            'PTNO
            'PATNAME
            'SEX
            'AGE
            'DPCD
            'WD_NO
            get_WEB_INTERFACE_S06 = ConvertXmlToRecordSet(strReturn)

        Catch ex As Exception
            get_WEB_INTERFACE_S06 = Nothing

            ErrMsgProc(strDiv & vbNewLine & strParam)

        Finally
            If Not objLIS Is Nothing Then objLIS.Dispose()

            objLIS = Nothing
        End Try
    End Function

    '검체번호, 검사코드로 결과입력
    Public Function insert_WEB_INTERFACE_I01(ByVal strSpcNo As String, ByVal strTestCd As String, ByVal strRstVal As String) As Boolean
        Dim strDiv As String
        Dim strParam As String
        Dim objLIS As webLIS.LisInterface
        Dim strReturn As String
        Dim i As Integer
        Dim aryTestCd() As String
        Dim aryRstVal() As String
        Dim strAllSpcNo As String
        Dim strAllTestCd As String
        Dim strAllRstVal As String
        Dim strAllErrFlag As String
        Dim strAllEquipCd As String
        Dim strAllGubun As String
        Dim strAllStsCd As String

        aryTestCd = Split(strTestCd, DLM_HS)
        aryRstVal = Split(strRstVal, DLM_HS)

        strAllSpcNo = ""
        strAllTestCd = ""
        strAllRstVal = ""
        strAllErrFlag = ""
        strAllEquipCd = ""
        strAllGubun = ""
        strAllStsCd = ""

        For i = LBound(aryTestCd) To UBound(aryTestCd)
            strAllSpcNo = strAllSpcNo & vbTab & strSpcNo
            strAllTestCd = strAllTestCd & vbTab & aryTestCd(i)
            strAllRstVal = strAllRstVal & vbTab & aryRstVal(i)
            strAllErrFlag = strAllErrFlag & vbTab
            strAllEquipCd = strAllEquipCd & vbTab & INS_CODE
            strAllGubun = strAllGubun & vbTab
            strAllStsCd = strAllStsCd & vbTab & "2"
        Next

        strAllSpcNo = strAllSpcNo & vbTab
        strAllTestCd = strAllTestCd & vbTab
        strAllRstVal = strAllRstVal & vbTab
        strAllErrFlag = strAllErrFlag & vbTab
        strAllEquipCd = strAllEquipCd & vbTab
        strAllGubun = strAllGubun & vbTab
        strAllStsCd = strAllStsCd & vbTab

        strDiv = "PG_SRL.INTERFACE_I01"

        strParam = "<Table>" & _
                          "<QID><![CDATA[PG_SRL.INTERFACE_I01]]></QID>" & _
                          "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                          "<USERID><![CDATA[" & gstrConnectTyp & "]]></USERID>" & _
                          "<EXECTYPE><![CDATA[NONQUERY]]></EXECTYPE>" & _
                          "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                          "<P0><![CDATA[" & strAllSpcNo & "]]></P0>" & _
                          "<P1><![CDATA[" & strAllTestCd & "]]></P1>" & _
                          "<P2><![CDATA[" & strAllRstVal & "]]></P2>" & _
                          "<P3><![CDATA[" & strAllErrFlag & "]]></P3>" & _
                          "<P4><![CDATA[" & strAllEquipCd & "]]></P4>" & _
                          "<P5><![CDATA[" & strAllGubun & "]]></P5>" & _
                          "<P6><![CDATA[" & UBound(aryTestCd) + 1 & "]]></P6>" & _
                          "<P7><![CDATA[" & "" & "]]></P7>" & _
                          "<P8><![CDATA[" & strAllStsCd & "]]></P8>" & _
                          "<P9><![CDATA[" & Space(1) & "]]></P9>" & _
                          "<P10><![CDATA[" & Space(4000) & "]]></P10>" & _
                   "</Table>"

        Try
            objLIS = New webLIS.LisInterface

            strReturn = objLIS.wsLISInterface(strDiv, strParam)

            insert_WEB_INTERFACE_I01 = ConvertXmlToBoolean(strReturn)

        Catch ex As Exception
            insert_WEB_INTERFACE_I01 = False

            ErrMsgProc(strDiv & vbNewLine & strParam)

        Finally
            If Not objLIS Is Nothing Then objLIS.Dispose()

            objLIS = Nothing
        End Try
    End Function
End Module
