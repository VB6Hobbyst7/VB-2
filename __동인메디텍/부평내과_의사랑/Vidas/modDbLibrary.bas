Attribute VB_Name = "modDbLibrary"
Option Explicit

Public XmlTxt As String
Public XmlTxtHead As String
Public XmlTxtTail As String
Public XMLAllTxt As String
Public XmlBody As String


Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim lsPid           As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String

    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))

        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
         
        SQL = ""
        SQL = "SELECT DISTINCT Company,HospCode,ChartNo,PatName,PatSex," & _
              "       PatAge,PatJumin,PatNo,CommDate,ExamNo," & _
              "       ExamID,ComExamID,Specimen,Result,Reference," & _
              "       Remark,RsltDate,IOFlag,TransYN,TransDT,Barcode,examtype " & vbCrLf & _
              "  FROM PAT_RES " & vbCrLf & _
              " WHERE EXAMTYPE = 'I' " & vbCrLf & _
              "   AND BARCODE = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' "
        'SQL = SQL & "  AND TRANSYN <> '2' "
        'SetRawData "[SQL]" & SQL
        
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult1 = ""
        sResult2 = ""
        
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 14)) '결과
            '-- 장비결과 치환
'            sResult1 = Replace(sResult1, "<", "")
'            sResult1 = Replace(sResult1, ">", "")
            
            If sResult1 <> "" Then
                Call SaveTrans(argSpcRow)
            End If
        Next iRow
        
        SaveTransDataW = 1
    
    End With

End Function


Private Sub SaveTrans(ByVal vasIDRow As Integer)
'선택전송
    'Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer
    Dim FindFile As String

'    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
'        Exit Sub
'    End If
    
    With frmInterface
        For vasResRow = 1 To .vasTemp.DataRowCnt
            .vasID.Row = vasIDRow
            .vasID.Col = 1
            If .vasID.Value = 1 Then
                .vasTemp.Row = vasResRow
                liRet = -1
                If Trim(GetText(.vasTemp, vasResRow, 3)) <> "" Then
                    liRet = Make_XML(vasResRow)
                End If
                
                If liRet = 1 Then
                    SetBackColor .vasID, vasIDRow, vasIDRow, colCheckBox, 12, 202, 255, 112
                    'SetText vasList, "전송완료", vasIDRow, colState
                    
                    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
                    If FindFile <> "" Then
                        '-- 2016.09.05
                        'Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     '전송완료가 됐을때 파일지우기
                    End If
                          
                          SQL = " Update pat_res Set "
                    SQL = SQL & " TransYN = '2', "
                    SQL = SQL & " TransDt = '" & Format(Now, "yyyymmdd") & "' "
                    .vasID.Row = vasIDRow: .vasID.Col = 4
                    SQL = SQL & " Where ChartNo  = '" & Trim(.vasID.Text) & "' "
                    .vasID.Row = vasIDRow: .vasID.Col = 12
                    SQL = SQL & "   and ExamID   = '" & Trim(.vasID.Text) & "' "
                    .vasID.Row = vasIDRow: .vasID.Col = 10
                    SQL = SQL & "   and CommDate = '" & Trim(.vasID.Text) & "'"
                    Res = SendQuery(gLocal, SQL)
                    
                Else
                    SetBackColor .vasID, vasIDRow, vasIDRow, colCheckBox, 12, 255, 0, 0
                    'SetText vasID, "실패", vasIDRow, colState
                End If
                '.vasID.Col = 1
                '.vasID.Value = "0"
            Else
            
            End If
        Next
    End With
    
    If XmlTxtHead = "" Then
        XmlTxtHead = "<?xml version=""1.0"" encoding=""euc-kr""?>" & vbCrLf & _
                     "<?xml-stylesheet type=""text/xsl"" href=C:\UBCare\SINAI\IF\Form\ExamIF_Form_05.xsl""?>" & vbCrLf & "<UBCare검사정보>"
    End If
    
    If XmlTxtTail = "" Then
        XmlTxtTail = "</UBCare검사정보>"
    End If
    
'    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
    SaveXMLFile XMLAllTxt
    
End Sub


Public Sub SaveXMLFile(argSQL As String, Optional argFlag As Integer = 0)
'argSQL의 내용을 파일로 저장
    Dim FilNum, FilNum1
    Dim FindFile As String
    Dim TxtString1 As String
    Dim AllString1 As String
    Dim i As Long
    
    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out.xml")
    
    
    If FindFile <> "" Then
'        Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
        FilNum1 = FreeFile
        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Input As FilNum1
        Do While Not EOF(FilNum1)
            Input #FilNum1, TxtString1
            AllString1 = AllString1 & TxtString1
        Loop
        Close #FilNum1
        
        i = InStr(1, AllString1, "</UBCare검사정보>")
        XmlBody = Mid(AllString1, 1, i - 1)
        argSQL = XmlBody & argSQL & XmlTxtTail
        '-- 2016.09.05
        Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
    Else
        argSQL = XmlTxtHead & argSQL & XmlTxtTail
    End If
    
'    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
    
    FilNum = FreeFile
    
    
    If argFlag = 0 Then
        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Output As FilNum
    Else
        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Append As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
    argSQL = ""
End Sub


Function Make_XML(asRow) As Integer
Dim varTmp As Variant
Dim strTmp As String
Dim strRslt As String

    With frmInterface.vasTemp
        .Row = asRow
                    XMLAllTxt = XMLAllTxt & "<검사>"
        .Col = 1:   XMLAllTxt = XMLAllTxt & "<업체>" & Trim(.Text) & "</업체>"
        .Col = 2:   XMLAllTxt = XMLAllTxt & "<요양기관번호>" & Trim(.Text) & "</요양기관번호>"
        .Col = 3:   XMLAllTxt = XMLAllTxt & "<차트번호>" & Trim(.Text) & "</차트번호>"
        .Col = 4:   XMLAllTxt = XMLAllTxt & "<수진자명>" & Trim(.Text) & "</수진자명>"
        .Col = 7:   XMLAllTxt = XMLAllTxt & "<주민등록번호>" & Trim(.Text) & "</주민등록번호>"
        .Col = 8:   XMLAllTxt = XMLAllTxt & "<내원번호>" & Trim(.Text) & "</내원번호>"
        .Col = 9:   XMLAllTxt = XMLAllTxt & "<의뢰일>" & Trim(.Text) & "</의뢰일>"
        .Col = 10:  XMLAllTxt = XMLAllTxt & "<검사번호>" & Trim(.Text) & "</검사번호>"
        .Col = 11:  XMLAllTxt = XMLAllTxt & "<검사ID>" & Trim(.Text) & "</검사ID>"
        .Col = 12:  XMLAllTxt = XMLAllTxt & "<업체검사ID></업체검사ID>"
        .Col = 13:  XMLAllTxt = XMLAllTxt & "<검체></검체>"
        .Col = 14:  strRslt = Trim(.Text)
                    XMLAllTxt = XMLAllTxt & "<결과치>" & strRslt & "</결과치>"
        .Col = 15:  XMLAllTxt = XMLAllTxt & "<참조치>" & Trim(.Text) & "</참조치>"
        .Col = 16:  XMLAllTxt = XMLAllTxt & "<소견>" & Trim(.Text) & "</소견>"
        .Col = 17:  XMLAllTxt = XMLAllTxt & "<결과일>" & Trim(.Text) & "</결과일>"
        .Col = 18:  XMLAllTxt = XMLAllTxt & "<입원외래구분>" & Trim(.Text) & "</입원외래구분>"
                    XMLAllTxt = XMLAllTxt & "</검사>"

    End With
    
    Make_XML = 1
    
End Function

Function SaveTransDataR(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow            As Integer
    Dim lsID            As String
    Dim lsPid           As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String

    With frmInterface
        SaveTransDataR = -1
        
        lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, colBarcode))
        lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, colPID))
        
        'Local에서 환자별로 결과값 가져오기
        ClearSpread frmInterface.vasTemp
        
        SQL = ""
        SQL = "SELECT DISTINCT Company,HospCode,ChartNo,PatName,PatSex," & _
              "       PatAge,PatJumin,PatNo,CommDate,ExamNo," & _
              "       ExamID,ComExamID,Specimen,Result,Reference," & _
              "       Remark,RsltDate,IOFlag,TransYN,TransDT,Barcode,examtype " & vbCrLf & _
              "  FROM PAT_RES " & vbCrLf & _
              " WHERE EXAMTYPE = 'I' " & vbCrLf & _
              "   AND BARCODE = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' "
              
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
        
        sResult1 = ""
        sResult2 = ""
        
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 14)) '결과
            '-- 장비결과 치환
            'sResult1 = Replace(sResult1, "<", "")
            'sResult1 = Replace(sResult1, ">", "")
            
            If sResult1 <> "" Then
                Call SaveTransR(argSpcRow)
            End If
        Next iRow
        
            
    End With
           
    SaveTransDataR = 1
    
End Function


Private Sub SaveTransR(ByVal vasIDRow As Integer)
'선택전송
    'Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer
    Dim FindFile As String

    
    With frmInterface
        For vasResRow = 1 To .vasTemp.DataRowCnt
            .vasRID.Row = vasIDRow
            .vasRID.Col = 1
            If .vasRID.Value = 1 Then
                .vasTemp.Row = vasResRow
                liRet = -1
                If Trim(GetText(.vasTemp, vasResRow, 3)) <> "" Then
                    liRet = Make_XML(vasResRow)
                End If
                
                If liRet = 1 Then
                    SetBackColor .vasRID, vasIDRow, vasIDRow, colCheckBox, 12, 202, 255, 112
                    'SetText vasList, "전송완료", vasIDRow, colState
                    
                    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
                    If FindFile <> "" Then
                        '-- 2016.09.05
                        'Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     '전송완료가 됐을때 파일지우기
                    End If
                          
                          SQL = " Update pat_res Set "
                    SQL = SQL & " TransYN = '2', "
                    SQL = SQL & " TransDt = '" & Format(Now, "yyyymmdd") & "' "
                    .vasRID.Row = vasIDRow: .vasRID.Col = 4
                    SQL = SQL & " Where ChartNo  = '" & Trim(.vasRID.Text) & "' "
                    .vasRID.Row = vasIDRow: .vasRID.Col = 12
                    SQL = SQL & "   and ExamID   = '" & Trim(.vasRID.Text) & "' "
                    .vasRID.Row = vasIDRow: .vasRID.Col = 10
                    SQL = SQL & "   and CommDate = '" & Trim(.vasRID.Text) & "'"
                    Res = SendQuery(gLocal, SQL)
                    
                Else
                    SetBackColor .vasRID, vasIDRow, vasIDRow, colCheckBox, 12, 255, 0, 0
                    'SetText vasID, "실패", vasIDRow, colState
                End If
                '.vasID.Col = 1
                '.vasID.Value = "0"
            Else
            
            End If
        Next
    End With
    
    If XmlTxtHead = "" Then
        XmlTxtHead = "<?xml version=""1.0"" encoding=""euc-kr""?>" & vbCrLf & _
                     "<?xml-stylesheet type=""text/xsl"" href=C:\UBCare\SINAI\IF\Form\ExamIF_Form_05.xsl""?>" & vbCrLf & "<UBCare검사정보>"
    End If
    
    If XmlTxtTail = "" Then
        XmlTxtTail = "</UBCare검사정보>"
    End If
    
'    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
    SaveXMLFile XMLAllTxt
    
End Sub

'-- 수진자 정보 가져오기
Function GetSampleInfoW(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim strAge  As String
    
    GetSampleInfoW = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '2 샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '바코드번호로 환자정보 불러오기
          SQL = "SELECT DiSTINCT CHARTNO, PATNAME, PATSEX, PATAGE,COMPANY,HOSPCODE,PATJUMIN,PATNO,COMMDATE,EXAMNO,EXAMID,IOFLAG  "
    SQL = SQL & vbCrLf & "  FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE EXAMTYPE = 'I' "
    SQL = SQL & vbCrLf & "   AND BARCODE = '" & sBarcode & "'"
    'SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
    

    Res = GetDBSelectColumn(gLocal, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID    '5
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName  '6
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '7
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colAge    '8
        
        SetText frmInterface.vasID, Trim(gReadBuf(4)), asRow, 12
        SetText frmInterface.vasID, Trim(gReadBuf(5)), asRow, 13
        SetText frmInterface.vasID, Trim(gReadBuf(6)), asRow, 14
        SetText frmInterface.vasID, Trim(gReadBuf(7)), asRow, 15
        SetText frmInterface.vasID, Trim(gReadBuf(8)), asRow, 16
        SetText frmInterface.vasID, Trim(gReadBuf(9)), asRow, 17
        SetText frmInterface.vasID, Trim(gReadBuf(10)), asRow, 18
        SetText frmInterface.vasID, Trim(gReadBuf(11)), asRow, 19
        
        GetSampleInfoW = 1
    Else
        GetSampleInfoW = -1
    End If

End Function

'-- 워크리스트 가져오기
Function GetSampleList(ByVal pFrDT As String, ByVal pToDT As String) As Integer
    
    Dim i   As Integer
    Dim RS As ADODB.Recordset
    Dim asRow As Integer
    
    GetSampleList = -1
    
    
    '바코드번호로 환자정보 불러오기
          SQL = "SELECT DiSTINCT CHARTNO, PATNAME, PATSEX, PATAGE,COMPANY,HOSPCODE,PATJUMIN,PATNO,BARCODE  "
    SQL = SQL & vbCrLf & "  FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE EXAMTYPE = 'I' "
    SQL = SQL & vbCrLf & "   AND COMMDATE BETWEEN '" & pFrDT & "' AND '" & pToDT & "'"
    SQL = SQL & vbCrLf & "   AND (RESULT = '' OR RESULT IS NULL) "
    SQL = SQL & " ORDER BY CHARTNO"
    
    Set RS = cn.Execute(SQL, , adCmdText)
    asRow = 0
    
    Do While Not RS.EOF
        asRow = asRow + 1
        frmInterface.vasID.MaxRows = asRow
        SetText frmInterface.vasID, Trim(RS.Fields(8).Value & ""), asRow, colBarcode
        SetText frmInterface.vasID, Trim(RS.Fields(0).Value & ""), asRow, colPID    '5
        SetText frmInterface.vasID, Trim(RS.Fields(1).Value & ""), asRow, colPName  '6
        SetText frmInterface.vasID, Trim(RS.Fields(2).Value & ""), asRow, colSex    '7
        SetText frmInterface.vasID, Trim(RS.Fields(3).Value & ""), asRow, colAge    '8

'''        SetText frmInterface.vasID, Trim(RS.Fields(4).Value & ""), asRow, 12
'''        SetText frmInterface.vasID, Trim(RS.Fields(5).Value & ""), asRow, 13
'''        SetText frmInterface.vasID, Trim(RS.Fields(6).Value & ""), asRow, 14
'''        SetText frmInterface.vasID, Trim(RS.Fields(7).Value & ""), asRow, 15
        
        
        SetText frmInterface.vasID, GetOrderCnt(Trim(RS.Fields(8).Value & ""), asRow), asRow, colCount
        SetText frmInterface.vasID, "0", asRow, colSndCnt
        SetText frmInterface.vasID, "0", asRow, colRcvCnt
        
        RS.MoveNext
        GetSampleList = 1
    Loop
    
    frmInterface.vasID.Col = 1
    frmInterface.vasID.Row = -1
    If frmInterface.vasID.Value = 0 Then
        frmInterface.vasID.Value = 1
    Else
        frmInterface.vasID.Value = 0
    End If
    RS.Close


End Function


'-- 검사갯수 가져오기
Function GetOrderCnt(ByVal pBARCODE As String, ByVal pRow As Integer) As Integer
    
    Dim i   As Integer
    Dim RS As ADODB.Recordset
    'Dim asRow As Integer
    Dim strIntBase As String
    Dim strTemp    As String
    
    GetOrderCnt = -1
    i = 0
    
    
    '바코드번호로 환자정보 불러오기
          SQL = "SELECT EXAMID  "
    SQL = SQL & vbCrLf & "  FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE EXAMTYPE = 'I' "
    SQL = SQL & vbCrLf & "   AND BARCODE = '" & pBARCODE & "' "
    SQL = SQL & vbCrLf & "   AND (RESULT = '' OR RESULT IS NULL) "
    SQL = SQL & vbCrLf & " ORDER BY COMEXAMID "
    
    Set RS = cn.Execute(SQL, , adCmdText)
    
    Do While Not RS.EOF
        'strIntBase = UCase(Trim(RS.Fields("EXAMID").Value & ""))
        
        If UCase(Trim(RS.Fields("EXAMID").Value & "")) = "OA" Then
            strIntBase = "FT4"
        ElseIf UCase(Trim(RS.Fields("EXAMID").Value & "")) = "AK" Then
            strIntBase = "TSH"
        Else
            strIntBase = UCase(Trim(RS.Fields("EXAMID").Value & ""))
        End If
        
        If strIntBase <> strTemp Then
            If frmInterface.vasID.MaxCols < colIntBase + i Then
                frmInterface.vasID.MaxCols = frmInterface.vasID.MaxCols + 1
            End If
            SetText frmInterface.vasID, strIntBase, pRow, colIntBase + i
            i = i + 1
        End If
        strTemp = strIntBase
        RS.MoveNext
    Loop
    
    RS.Close

    GetOrderCnt = i

End Function

Function GetSampleInfoR(ByVal asRow As Long) As Integer
    Dim sBarcode As String
    Dim sSpecNo As String

    GetSampleInfoR = -1
    
    '환자정보 가져오기
    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBarcode))   '샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '바코드번호로 환자정보 불러오기

    SQL = ""
    SQL = SQL + "SELECT " & gDBCOLUMN_Parm.PID & "," & gDBCOLUMN_Parm.PNAME & "," & gDBCOLUMN_Parm.PSEX & "," & gDBCOLUMN_Parm.PAGE + vbLf
    SQL = SQL + "  FROM " & gDBTBL_Parm.ORDTABLE + vbLf
    SQL = SQL + " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " + vbLf
    SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " + vbLf
    SQL = SQL + "   AND " & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL" + vbLf
    
    Res = GetDBSelectColumn(gServer, SQL)
    
    If Res = 1 Then
        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colAge
        
        GetSampleInfoR = 1
    Else
    
        GetSampleInfoR = -1
    End If
End Function

Function GetEquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String

    GetEquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    Res = GetDBSelectVas(gLocal, SQL, frmInterface.vasTemp1)
    
    If frmInterface.vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To frmInterface.vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
        End If
    Next i

    'SPSLHRRST
    SQL = " Select SUCD From LRESULT " & vbCr & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    Res = GetDBSelectColumn(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        GetEquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function


Function GetGetEquipExamCode_CA1500(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim strStatFg  As String
Dim sExamCd As String
Dim strItems As String
Dim strTemp As String
Dim strIntBase As String

    GetGetEquipExamCode_CA1500 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
'    argPID = "1558200030"
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    Res = GetDBSelectColumn(gServer, SQL)
    GetGetEquipExamCode_CA1500 = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            sExamCd = Trim(gReadBuf(i))
        Else
            Exit For
        End If
    Next
    
    SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and SUBSTR(exmn_cd,1,1) <> 'G'" & _
          "   and RSLT_NO IS NOT NULL"
          
    Res = GetDBSelectRow(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    Erase gReadBuf
          SQL = "Select equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    SQL = SQL & " order by equipcode    "
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strIntBase = Trim(gReadBuf(i))
            strIntBase = Mid(strIntBase, 1, Len(strIntBase) - 1) & "0" & Space$(6)
            If strIntBase <> strTemp Then
                strExamCode = strExamCode & strIntBase 'Mid(Trim(gReadBuf(i)), 1, Len(Trim(gReadBuf(i))) - 1) & "0" & Space$(6)
                strTemp = strIntBase
            End If

            'strExamCode = strExamCode & Mid(Trim(gReadBuf(i)), 1, Len(Trim(gReadBuf(i))) - 1) & "0" & Space$(6)
        Else
            Exit For
        End If
    Next
    
    '응급유무 (R:Routin, E:Stat)
    'strStatFg = IIf(pAccInfo.StatFg = "1", "E", "U")
    strStatFg = "U"
    
    
'    strExamCode = STX & "S2210101" & strStatFg & Space(6) & Space(4) & mOrder.RackNo & mOrder.TubePos & mOrder.BarNo & _
                "B" & Space(15) & strExamCode & ETX
    
    strExamCode = "" & "S2210101" & strStatFg & Space(6) & Space(4) & mResult.RackNo & mResult.TubePos & mResult.BarNo & _
                "B" & Space(15) & strExamCode & ""
    
    GetGetEquipExamCode_CA1500 = strExamCode
    
End Function

Function GetOrderExamCode(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String

    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    
    '-- 검사코드 가져오기
          SQL = "SELECT DiSTINCT b.SCP42SUGACD "
    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & argPID & "'"
    SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
          
    Res = GetDBSelectColumn(gServer, SQL)
    GetOrderExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
    End If
    
End Function

Function GetOrderExamCode_New(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String
Dim rs_svr As ADODB.Recordset

    GetOrderExamCode_New = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    argPID = Mid(argPID, 1, 14)
    
          SQL = "SELECT DiSTINCT EXAMID "
    SQL = SQL & vbCrLf & "  FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE EXAMTYPE = 'I' "
    SQL = SQL & vbCrLf & "   AND BARCODE = '" & argPID & "'"
    
    Set rs_svr = cn.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_New = GetOrderExamCode_New & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
    If GetOrderExamCode_New <> "" Then
        GetOrderExamCode_New = Mid(GetOrderExamCode_New, 1, Len(GetOrderExamCode_New) - 1)
    End If
    
End Function


Function GetGetEquipExamCode_E411(argEquipCode As String, argPID As String, Optional intRow As Long) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim sSpecNo     As String
Dim iRow        As Long
Dim SpecNo      As String
    
    GetGetEquipExamCode_E411 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '-- 자검체는 11자리임 조회하기위하여 마지막 자리를 없앤다.
    argPID = Mid(argPID, 1, 10)
    
    If Mid(argPID, 1, 2) = "99" Then
        'strExamCode = Proc_Order_LX_QC(argPID)
        
        'iRow = frmInterface.vasID.DataRowCnt
        iRow = intRow
        
        SpecNo = Trim(GetText(frmInterface.vasID, iRow, colSpecNo))
        
        SQL = "SELECT QC_EXMN_CD "
        SQL = SQL & vbCrLf & " FROM SPSLMQMST "
        SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(SpecNo, 3, 3) & "' "     '//// 장비 번호
        SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(SpecNo, 6, 3) & "' "     '//// 검사명 번호
        SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(SpecNo, 9, 1) & "' "      '//// 레벨 번호
        SQL = SQL & vbCrLf & "  AND QC_EXMN_CD IN (" & gAllExam & ") "
        SQL = SQL & vbCrLf & "  AND USE_STR_DT <= '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' "
        SQL = SQL & vbCrLf & "  AND USE_END_DT >= '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' "
        Res = GetDBSelectRow(gServer, SQL)
        strExamCode = ""
        
        For i = 0 To UBound(gReadBuf)
            If gReadBuf(i) <> "" Then
                strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
            Else
                Exit For
            End If
        Next
        
    Else
        '바코드번호로 검체번호 불러오기
        SQL = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
        Res = GetDBSelectColumn(gServer, SQL)
        sSpecNo = Trim(gReadBuf(0))
        
        '-- 검사코드 가져오기
        SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
              " Where SPCM_NO = '" & Trim(sSpecNo) & "' " & vbCrLf & _
              "   and RSLT_NO IS NOT NULL"
              
        Res = GetDBSelectRow(gServer, SQL)
        strExamCode = ""
        
        For i = 0 To UBound(gReadBuf)
            If gReadBuf(i) <> "" Then
                strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
            Else
                Exit For
            End If
        Next
    End If
    
    If strExamCode = "" Then
'        MsgBox "미접수 환자"
        GetGetEquipExamCode_E411 = ""
        Exit Function
    End If
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select distinct equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
    
        If gReadBuf(i) <> "" Then
            'gReadBuf(i) = Mid(gReadBuf(i), 1, Len(gReadBuf(i)) - 1)
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next
    
    GetGetEquipExamCode_E411 = Mid(strExamCode, 2)
    
End Function



Function GetGetEquipExamCode_Architect(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_Architect = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- 검사코드 가져오기
'    SQL = ""
'    SQL = SQL + "SELECT " & gDBCOLUMN_Parm.TESTCD & vbCrLf
'    SQL = SQL + "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
'    SQL = SQL + " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
'    SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
'    SQL = SQL + "   AND " & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL"
'
'    Res = GetDBSelectRow(gServer, SQL)
'    strExamCode = ""
'
'    For i = 0 To UBound(gReadBuf)
'        If gReadBuf(i) <> "" Then
'            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
'        Else
'            Exit For
'        End If
'    Next
'
'    If strExamCode = "" Then
'        '-- 미접수환자이거나 해당장비에 검사대상 없음
'        GetGetEquipExamCode_Architect = ""
'        Exit Function
'    End If
'
'    '-- 마지막 "," 자르기
'    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    
    ClearSpread frmInterface.vasTemp1
    
    '-- 가져온 검사코드의 채널 찾기
    SQL = "          "
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQUIPEXAM "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    
    '-- 해당 장비에 맞게 오더채널 가공하기 [ASTM Format >> Architect]
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next
    
    '-- 첫자리 "\" 자르기
    GetGetEquipExamCode_Architect = strExamCode
    
End Function


Function GetGetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_AU480 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    
    ClearSpread frmInterface.vasTemp1
    
    '-- 가져온 검사코드의 채널 찾기
    SQL = "          "
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQUIPEXAM "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            'If Trim(gReadBuf(i)) <> "990" Then
                '                                                     dilution
                strExamCode = strExamCode & "0" & Trim(gReadBuf(i)) & "0"
            'End If
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_AU480 = strExamCode
    
End Function


Function GetGetEquipExamCode(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String

    GetGetEquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '-- 자검체는 11자리임 조회하기위하여 마지막 자리를 없앤다.
    argPID = Mid(argPID, 1, 10)
    
    SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
          " Where SPCM_NO = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
          
    Res = GetDBSelectColumn(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
          SQL = "Select equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(argEquipCode) & ")"
    
    Res = GetDBSelectColumn(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & Trim(gReadBuf(i)) & "0" & Space$(6)
        Else
            Exit For
        End If
    Next
    
    GetGetEquipExamCode = strExamCode
    
End Function


