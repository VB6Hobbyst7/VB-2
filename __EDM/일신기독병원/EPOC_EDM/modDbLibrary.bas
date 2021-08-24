Attribute VB_Name = "modDbLibrary"
Option Explicit


Public Function db_select_Vas(argServer As Integer, argSQL As String, ByVal argSpread As vaSpread, Optional argRow As Integer = 1, Optional argcol As Integer = 1) As Integer
'쿼리 실행 내용을 스프레드쉬트에 Display
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Vas = -1
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set RS = cmdSQL.Execute
  
    If argSpread.MaxCols < RS.Fields.Count + argcol - 1 Then
        argSpread.MaxCols = RS.Fields.Count + argcol - 1
    End If
    
    If RS.EOF = True Or RS.BOF = True Then
        db_select_Vas = 0
        Exit Function
    End If
    
    'rs.MoveFirst
    i = argRow
    While Not RS.EOF
        If argSpread.MaxRows < i Then
            argSpread.MaxRows = i
        End If
        
        For j = 0 To RS.Fields.Count - 1
            argSpread.Row = i
            argSpread.Col = j + argcol
            If IsNull(RS.Fields.Item(j).Value) Then
                argSpread.Text = ""
            Else
                argSpread.Text = Trim(CStr(RS.Fields.Item(j).Value))
            End If
        Next j
        i = i + 1
        RS.MoveNext
    Wend
    
    
    
    If argSpread.DataRowCnt = 0 Then
        db_select_Vas = 0
    Else
        db_select_Vas = i - 1
        'argSpread.MaxRows = i - 1
    End If
    
    RS.Close
    
    Exit Function
ErrHandle:
'    MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Vas = -1
    
End Function

Public Function db_select_Col(argServer As Integer, argSQL As String) As Integer
'쿼리 실행 내용을 gReadbuf()의 Array에 저장
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Col = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set RS = cmdSQL.Execute
           
    If Not (RS.EOF Or RS.BOF) Then
        'rs.MoveFirst
    Else
        db_select_Col = 0
        gReadBuf(0) = ""
        RS.Close
        Exit Function
    End If
    
    
    Do While Not RS.EOF
        For i = 0 To RS.Fields.Count - 1
            If IsNull(RS.Fields.Item(i).Value) = True Then
                gReadBuf(i) = ""
            Else
                gReadBuf(i) = Trim(CStr(RS.Fields.Item(i).Value))
            End If
        Next i
        
        db_select_Col = 1
        
        RS.MoveNext
        Exit Do
    Loop
    
    RS.Close
    
    Exit Function
    
ErrHandle:
    'MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Col = -1
End Function

Private Function f_subSet_RefVal(ByVal strORCD As String, ByVal strSubCD As String, Optional ByVal strRslt As String, Optional ByVal strSex As String, Optional ByVal strAge As String) As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
Dim rs_svr As ADODB.Recordset

On Error GoTo ErrorTrap
    
    strRslt = Replace(strRslt, "<", "")
    strRslt = Replace(strRslt, ">", "")
    f_subSet_RefVal = " "
    
    f_subSet_RefVal = ""
          SQL = "Select REFHIGH, REFLOW  "
    SQL = SQL & "  From EQPMASTER"
    SQL = SQL & " Where EQUIPNO = '" & gEquip & "' "
    SQL = SQL & "   And EXAMCODE =  '" & strORCD & "'"
'    SQL = SQL & "   And SUBCODE =  '" & strSubCD & "'"
    
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If Res > 0 Then
        If IsNumeric(strRslt) And IsNumeric(Trim(gReadBuf(0))) And IsNumeric(Trim(gReadBuf(1))) Then
            If Val(strRslt) > Val(Trim(gReadBuf(0))) Then
                f_subSet_RefVal = "H"
            ElseIf Val(strRslt) < Val(Trim(gReadBuf(1))) Then
                f_subSet_RefVal = "L"
            Else
                f_subSet_RefVal = " "
            End If
        Else
            f_subSet_RefVal = " "
        End If
    End If
    
Exit Function

ErrorTrap:
    f_subSet_RefVal = " "
'    Set AdoRs_ORACLE = Nothing
'    Call ErrMsgProc(CallForm)
     
End Function



Public Sub SaveTrans(ByVal vasIDRow As Integer)
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
                    SetBackColor .vasID, vasIDRow, vasIDRow, colCHECKBOX, colState - 1, 202, 255, 112
                    SetText .vasID, "전송완료", vasIDRow, colState
                    
                    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
                    If FindFile <> "" Then
                        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     '전송완료가 됐을때 파일지우기
                    End If
                          
                          SQL = " Update PATRESULT Set "
                    SQL = SQL & " SENDFLAG = '2', "
                    SQL = SQL & " SENDDATE = '" & Format(Now, "yyyymmdd") & "' "
                    SQL = SQL & " Where BARCODE  = '" & GetText(frmInterface.vasID, vasIDRow, colBARCODE) & "' "
                    SQL = SQL & "   AND SAVESEQ  = " & GetText(frmInterface.vasID, vasIDRow, colSAVESEQ)
                    SQL = SQL & "   AND MID(EXAMDATE,1,8)   = '" & Mid(GetText(frmInterface.vasID, vasIDRow, colEXAMDATE), 1, 8) & "' "
                    Res = SendQuery(gLocal, SQL)

                Else
                    SetBackColor .vasID, vasIDRow, vasIDRow, colCHECKBOX, 12, 255, 0, 0
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
        'Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
        FilNum1 = FreeFile
        Open "C:\UBCare\SINAI\IF\ExamIF_out.xml" For Input As FilNum1
        
        Do While Not EOF(FilNum1)
            Input #FilNum1, TxtString1
            Line Input #FilNum1, TxtString1
            AllString1 = AllString1 & TxtString1
        Loop

        Close #FilNum1
        i = InStr(1, AllString1, "</UBCare검사정보>")
        XmlBody = Mid(AllString1, 1, i - 1)
        argSQL = XmlBody & argSQL & XmlTxtTail
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
'Dim varTmp As Variant
'Dim strTmp As String
'Dim strRslt As String

'    With frmInterface.vasTemp
'        .Row = asRow
'                    XMLAllTxt = XMLAllTxt & "<검사>"
'        .Col = 1:   XMLAllTxt = XMLAllTxt & "<업체>" & Trim(.Text) & "</업체>"
'        .Col = 2:   XMLAllTxt = XMLAllTxt & "<요양기관번호>" & Trim(.Text) & "</요양기관번호>"
'        .Col = 3:   XMLAllTxt = XMLAllTxt & "<차트번호>" & Trim(.Text) & "</차트번호>"
'        .Col = 4:   XMLAllTxt = XMLAllTxt & "<수진자명>" & Trim(.Text) & "</수진자명>"
'        .Col = 7:   XMLAllTxt = XMLAllTxt & "<주민등록번호>" & Trim(.Text) & "</주민등록번호>"
'        .Col = 8:   XMLAllTxt = XMLAllTxt & "<내원번호>" & Trim(.Text) & "</내원번호>"
'        .Col = 9:   XMLAllTxt = XMLAllTxt & "<의뢰일>" & Trim(.Text) & "</의뢰일>"
'        .Col = 10:  XMLAllTxt = XMLAllTxt & "<검사번호>" & Trim(.Text) & "</검사번호>"
'        .Col = 11:  XMLAllTxt = XMLAllTxt & "<검사ID>" & Trim(.Text) & "</검사ID>"
'        .Col = 12:  XMLAllTxt = XMLAllTxt & "<업체검사ID></업체검사ID>"
'        .Col = 13:  XMLAllTxt = XMLAllTxt & "<검체></검체>"
'        .Col = 14:  strRslt = Trim(.Text)
'                    XMLAllTxt = XMLAllTxt & "<결과치>" & strRslt & "</결과치>"
'        .Col = 15:  XMLAllTxt = XMLAllTxt & "<참조치>" & Trim(.Text) & "</참조치>"
'        .Col = 16:  XMLAllTxt = XMLAllTxt & "<소견>" & Trim(.Text) & "</소견>"
'        .Col = 17:  XMLAllTxt = XMLAllTxt & "<결과일>" & Trim(.Text) & "</결과일>"
'        .Col = 18:  XMLAllTxt = XMLAllTxt & "<입원외래구분>" & Trim(.Text) & "</입원외래구분>"
'                    XMLAllTxt = XMLAllTxt & "</검사>"
'
'    End With
'
'    Make_XML = 1
    Dim FilNum
    Dim FilNum2
    Dim TxtString As String
    Dim ResultString As String
    Dim TxtRece As String
    Dim i As Long
    Dim PChartNum As String
    Dim PName As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAge As String
    Dim pSex As String
    Dim STxt, NumTxt As Long
    Dim SQL As String
    Dim PEquipno As String
    
    Dim PExamname As String
    Dim PEquipCode As String
    Dim j As Long
    Dim BarFlag As Integer
    Dim pResult As String
    Dim pExamdate As String
    Dim pOpinion As String
    Dim TxtPat As String
    Dim IOGubun As String
    Dim TestNum As String
    
    Make_XML = -1
    
    ClearSpread frmInterface.vasResTemp
    
    SQL = "select  chartno, examcode, hospdate, barcode,pname, pjumin, examdate, '', seqno,result " & vbCrLf & _
          "  from PATRESULT " & vbCrLf & _
          " where mid(examdate,1,8) = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' " & vbCrLf & _
          "   and result <> '' " & vbCrLf & _
          "   And equipno = '" & gEquip & "' " & _
          "   and barcode = '" & Trim(GetText(frmInterface.vasID, asRow, 5)) & "'"
    SQL = SQL & "order by seqno "
    Res = db_select_Vas(gLocal, SQL, frmInterface.vasResTemp)

    For i = 1 To frmInterface.vasResTemp.DataRowCnt
        PID = Trim(GetText(frmInterface.vasResTemp, i, 1))
        PExamCode = Trim(GetText(frmInterface.vasResTemp, i, 2))
        PReceDate = Trim(GetText(frmInterface.vasResTemp, i, 3))
        PChartNum = Trim(GetText(frmInterface.vasResTemp, i, 4))
        PName = Trim(GetText(frmInterface.vasResTemp, i, 5))
        PJumin = Mid(Trim(GetText(frmInterface.vasResTemp, i, 6)), 1, 6) & "-" & Mid(Trim(GetText(frmInterface.vasResTemp, i, 6)), 7)
        pExamdate = Trim(GetText(frmInterface.vasResTemp, i, 7))
        IOGubun = Trim(GetText(frmInterface.vasResTemp, i, 8))
        TestNum = Trim(GetText(frmInterface.vasResTemp, i, 9))
        pResult = Trim(GetText(frmInterface.vasResTemp, i, 10))
        XMLAllTxt = XMLAllTxt & "<검사><업체>ACK</업체><요양기관번호>38341948</요양기관번호><차트번호>" & PChartNum & "</차트번호><수진자명>" & PName & "</수진자명><주민등록번호>" & PJumin & "</주민등록번호><내원번호>" & PID & "</내원번호><의뢰일>" & PReceDate & "</의뢰일><검사번호>" & TestNum & "</검사번호><검사ID>" & PExamCode & "</검사ID><업체검사ID></업체검사ID><검체></검체><결과치>" & pResult & "</결과치><참조치></참조치><소견></소견><결과일>" & pExamdate & "</결과일><입원외래구분>" & IOGubun & "</입원외래구분></검사>"
    Next
    
    Make_XML = 1
    
End Function

Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim strDate         As String
    Dim strInNum        As String
    Dim strGumNum       As String
    Dim VallsID         As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim strSubCD        As String
    Dim strRefVal       As String
    Dim strSpcCd        As String
    Dim strSex As String
    Dim strAge  As String
    Dim strORQN As String
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    
    Dim strReceNo   As String
    Dim strSeqNo   As String
    
    Dim tmpREF As String
    Dim strREF As String
    Dim GumEqpCd As String * 100
    
    Dim strExamDate As String
    
    Dim strKey1     As String
    Dim strKey2     As String
    Dim strSaveSeq  As String
    Dim strSubCodes As String
    Dim strChtNum   As String
    Dim strChannel As String
    Dim strReturn  As String
    Dim strRsltType As String
    
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    Dim prm6 As New ADODB.Parameter
    Dim prm7 As New ADODB.Parameter
    
    
    Dim strLABMACRACKNO As String
    
'On Error GoTo ErrHandle

    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        If InStr(lsID, "오더없음") > 0 Then
            Exit Function
        End If
        
        If Len(lsID) < 11 Then
            Exit Function
        End If
        
        If Not IsNumeric(lsID) Then
            Exit Function
        End If
        
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strChtNum = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        strExamDate = Trim(GetText(.vasID, argSpcRow, colEXAMDATE))
        strSaveSeq = Trim(GetText(.vasID, argSpcRow, colSAVESEQ))
        
        strLABMACRACKNO = "E" & Format(strSaveSeq, "0000")
        
        '-- Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE,INOUT,EXAMNAME " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '장비코드
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'  " & vbCrLf                                      '검사일
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf       '바코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq       '저장번호
'        SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colDISKNO)) & "' " & vbCrLf         'DISK 번호(진료검사ID)
'        SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPOSNO)) & "' "                    'POS 번호(진료지원ID)
              
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult = ""
        sResult1 = ""
        sResult2 = ""
        strKey1 = ""
        strKey2 = ""
        
        cn_Ser.BeginTrans
        
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strChannel = Trim(GetText(.vasTemp, iRow, 1))
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '결과(수정결과)
'            strChannel = Trim(GetText(.vasTemp, iRow, 16))
'            strSex = Trim(GetText(.vasTemp, iRow, 8))
'            strAge = Trim(GetText(.vasTemp, iRow, 10))
'            strORQN = Trim(GetText(.vasTemp, iRow, 14))
            
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If sResult <> "" Then
                '-- 서버저장
                      SQL = "INSERT INTO MEDI.LMI_RESULT"
                SQL = SQL & "(SPECIMEN_SER, LAB_MAC_CODE, LMI_ITEM_CODE, EXAM_RESULT_SEQ, EXAM_RESULT, LAB_MAC_RACK_NO, LAB_MAC_TUBE_POSITION, STATUS_FLAG, RESULT_DATE) "
                SQL = SQL & " VALUES "
                SQL = SQL & "('" & lsID & "'"
                SQL = SQL & ", '" & gEquipCode & "'"
                SQL = SQL & ", '" & strChannel & "'"
                SQL = SQL & ", MEDI.FN_CPL_LOAD_MAX_LMI_RESULT('" & lsID & "','" & gEquipCode & "','" & strChannel & "')"
                SQL = SQL & ", '" & sResult & "'"
                SQL = SQL & ", '" & strLABMACRACKNO & "' "                     ' LAB_MAC_RACK_NO" '==>  E0001'
                SQL = SQL & ", '' "                       ' LAB_MAC_TUBE_POSITION"
                SQL = SQL & ", '0' "                      ' STATUS_FLAG (0:정상)"
                SQL = SQL & ", SYSDATE )"
                
                Call SetSQLData("결과저장", SQL)
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow
        
        cn_Ser.CommitTrans
        SaveTransDataW = 1
    
    
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
    cn_Ser.RollbackTrans
    
End Function


'Function SaveTransDataR(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
''서버의 데이타 베이스에 저장
'    Dim iRow            As Integer
'    Dim lsID            As String
'    Dim lsPid           As String
'    Dim sResult         As String
'    Dim sResult1        As String
'    Dim sResult2        As String
'    Dim strEqpCd        As String
'    Dim VallsID         As String
'    Dim strDate         As String
'
'    SaveTransDataR = -1
'
'    'Local에서 환자별로 결과값 가져오기
'    ClearSpread frmInterface.vasTemp
'
'    With frmInterface
'        lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, 2))
'        VallsID = lsID
'        lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, 5))
'        strDate = Format(CDate(.dtpExamDate.Value), "yyyymmdd")
'
'        '-- Local에서 환자별로 결과값 가져오기
'        ClearSpread .vasTemp
'
'              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX " & vbCrLf
'        SQL = SQL & "  FROM PATRESULT " & vbCrLf
'        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                            '장비코드
'        SQL = SQL & "   AND EXAMDATE = '" & strDate & "'  " & vbCrLf   '검사일
'        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.vasRID, argSpcRow, 2)) & "' " & vbCrLf     '바코드
'        'SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf         'DISK 번호
'        'SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "                    'POS 번호
'
'        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
'
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'
'        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
'
'        sResult = ""
'        sResult1 = ""
'        sResult2 = ""
'
'        cn_Ser.BeginTrans
'
'        '서버로 결과값 저장하기
'        For iRow = 1 To .vasTemp.DataRowCnt
'            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
'            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '결과(장비결과)
'            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '결과(수정결과)
'
'            '-- 장비결과적용
'            If .optSaveResultR(0).Value = True Then
'                sResult = sResult1
'            Else
'                sResult = sResult2
'            End If
'
'            If sResult <> "" Then
''                If Len(VallsID) > 6 Then
''                    SQL = "Update ONIT..GUMJIN_INTERFACE" & _
''                          "   Set RESULT = '" & sResult & "'," & _
''                          "       ACT_RETURN_DATE = '" & strDate & "'" & _
''                          " Where PER_GUMJIN_DATE = '" & Mid(lsID, 1, 8) & "'" & _
''                          "   And PER_GUM_NUM = " & lsID & "" & _
''                          "   And INTERFACECODE = '" & strEqpCd & "'"
''                Else
'                    SQL = "Update onit_out..jun370_resulttb" & _
'                          "   Set Result = '" & sResult & "'" & _
'                          " Where orderorder = '" & VallsID & "'" & _
'                          "   and map2seqno = '" & strEqpCd & "'"
''                End If
'
'                Res = SendQuery(gServer, SQL)
'
'                If Res < 0 Then
'                    SaveQuery SQL
'                    cn_Ser.RollbackTrans
'                    Exit Function
'                End If
'
'            End If
'        Next iRow
'
'    End With
'
'    cn_Ser.CommitTrans
'    SaveTransDataR = 1
'
'End Function

'-- 검사자 정보 가져오기
Function GetSampleInfoW(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    
    GetSampleInfoW = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'          SQL = " SELECT DISTINCT '' AS 접수일자"
'    SQL = SQL & ", '' AS 차트번호"
'    SQL = SQL & ", '' AS 내원번호"
'    SQL = SQL & ", '' AS 입외"
'    SQL = SQL & ", '' AS 이름"
'    SQL = SQL & ", '' AS 성별"
'    SQL = SQL & ", '' AS 나이" & vbCrLf
'    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
'    SQL = SQL & " WHERE QC_BAR_NO = '" & sBarcode & "'" & vbCrLf
    
    
                                   
     
          'SQL = "SELECT P.PbsPatNam, O.OSPCHTNUM, R.ResLabCod, E.LabShtNam, R.ResOcmNum, R.ResOdrSeq, R.ResSeq, R.ResSubSeq, R.ResRltVal" & vbCrLf
          SQL = "SELECT DISTINCT P.PbsPatNam, O.OSPCHTNUM, R.ResOcmNum " & vbCrLf
    SQL = SQL & "  FROM RsbInf M, ResInf R, ospinf O, PBSINF P, LabMst E" & vbCrLf
    SQL = SQL & " WHERE M.RsbBarCod = '" & sBarcode & "'" & vbCrLf
    SQL = SQL & "   And M.RsbAckStt <> 'A' " & vbCrLf
    SQL = SQL & "   And O.OspChkStt <> 'F' " & vbCrLf
    SQL = SQL & "   And (R.ResRepTyp Is Null or R.ResRepTyp <> 'F') " & vbCrLf
    SQL = SQL & "   And (R.ResRltVal is Null or R.ResRltVal = '') " & vbCrLf
    SQL = SQL & "   And R.ResStatus <> '5' " & vbCrLf
    SQL = SQL & "   And M.RSBACPNUM = R.ResRsbAcp" & vbCrLf
    SQL = SQL & "   And R.ResOcmNum = O.OspOcmNum" & vbCrLf
    SQL = SQL & "   and R.ResOdrSeq = O.OspOdrSeq" & vbCrLf
    SQL = SQL & "   and R.ResSeq    = O.OspSeq" & vbCrLf
    SQL = SQL & "   and O.OSPCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   and R.ResLabcod = E.LabCod" & vbCrLf
    
    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, "1", asRow, colCHECKBOX
        SetText frmInterface.vasID, sBarcode, asRow, colBARCODE
        'SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colHOSPDATE       '접수일
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colCHARTNO        '챠트번호
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID            '등록번호(저장시 필요)
        'SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colINOUT          '입/외
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPNAME          '환자명
        'SetText frmInterface.vasID, Trim(gReadBuf(5)), asRow, colPSEX           '성별
        'SetText frmInterface.vasID, Trim(gReadBuf(6)), asRow, colPAGE           '나이
        
        GetSampleInfoW = 1
   
    Else
        GetSampleInfoW = -1
    End If

    frmInterface.vasID.RowHeight(-1) = 12

End Function


'-- 검사자 정보 가져오기
Function GetSampleInfoW_JBUNIV(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    
    
    GetSampleInfoW_JBUNIV = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
'    strTestCd = mGetP(frmInterface.cboTest.Text, 2, "|")
    pFrDt = Format(frmInterface.dtpStartDt.Value, "yyyymmdd") & "000000"
    pToDt = Format(frmInterface.dtpStopDt.Value, "yyyymmdd") & "235959"
'    pFrNo = frmInterface.txtStartNum.Text
'    pToNo = frmInterface.txtStopNum.Text
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- 전북대병원  r010m.SPCCD
    SQL = ""
    SQL = SQL & "SELECT '1', '' AS SN ,'' AS 결과일시, j011m.colldt AS 접수일자, j011m.bcno AS 바코드번호, j010m.bcprtno AS 차트번호" & vbCr
    SQL = SQL & "       , r010m.WKYMD||r010m.WKGRPCD||r010m.WKNO FLWKNO " & vbCr
    SQL = SQL & "       , r010m.WKNO AS 접수번호" & vbCr
    SQL = SQL & "       , j011m.regno AS 내원번호" & vbCr
    SQL = SQL & "       , j010m.patnm AS 이름" & vbCr
    SQL = SQL & "       , j010m.age AS 나이" & vbCr
    SQL = SQL & "       , j010m.sex AS 성별" & vbCr
    SQL = SQL & "       , j011m.IOGBN" & vbCr
    SQL = SQL & "       , j010m.DEPTCD" & vbCr
    SQL = SQL & "       , j010m.WARDNO" & vbCr
    SQL = SQL & "       , j010m.ROOMNO" & vbCr
    SQL = SQL & "       , f72m.testcd AS ITEM" & vbCr
    SQL = SQL & "       , r010m.SPCCD AS SPCCD " & vbCr
    SQL = SQL & "  FROM LJ011M j011m                                     " & vbCr
    SQL = SQL & "       INNER JOIN LJ010M j010m                          " & vbCr
    SQL = SQL & "               ON j011m.bcno  = j010m.bcno              " & vbCr
    SQL = SQL & "              AND j011m.regno = j010m.regno             " & vbCr
    SQL = SQL & "       INNER JOIN LR010M r010m                          " & vbCr
    SQL = SQL & "               ON j011m.bcno   = r010m.bcno             " & vbCr
    SQL = SQL & "              AND j011m.regno  = r010m.regno            " & vbCr
    SQL = SQL & "              AND NVL(r010m.rstflg,'0') = '0'           " & vbCr
    SQL = SQL & "       INNER JOIN LF072M f72m                           " & vbCr
    SQL = SQL & "               ON f72m.eqcd    = 'G0006'                " & vbCr
    SQL = SQL & "              AND f72m.testcd  = '" & strTestCd & "'    " & vbCr
    SQL = SQL & "              AND r010m.testcd = f72m.testcd            " & vbCr
    SQL = SQL & " WHERE j011m.colldt BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
    SQL = SQL & "   AND r010m.wkno between '" & pFrNo & "' AND '" & pToNo & "' " & vbCr
    SQL = SQL & "   AND j011m.spcflg  = '4'                        " & vbCr
    SQL = SQL & "   AND NVL(j011m.rstflg, '0')  = '0'            " & vbCr
    SQL = SQL & "   AND j011m.bcno = '" & sBarcode & "'" & vbCr
    SQL = SQL & " UNION                                              " & vbCr
    SQL = SQL & "SELECT '1', '' AS SN ,'' AS 결과일시, j011m.colldt AS 접수일자, j011m.bcno AS 바코드번호, j010m.bcprtno AS 차트번호 " & vbCr
    SQL = SQL & "        , r010m.FLWKNO" & vbCr
    SQL = SQL & "        , r010m.WKNO AS 접수번호" & vbCr
    SQL = SQL & "        , j011m.regno AS 내원번호" & vbCr
    SQL = SQL & "        , j010m.patnm AS 이름" & vbCr
    SQL = SQL & "        , j010m.age AS 나이" & vbCr
    SQL = SQL & "        , j010m.sex AS 성별" & vbCr
    SQL = SQL & "        , j011m.IOGBN" & vbCr
    SQL = SQL & "        , j010m.DEPTCD" & vbCr
    SQL = SQL & "        , j010m.WARDNO" & vbCr
    SQL = SQL & "        , j010m.ROOMNO" & vbCr
    SQL = SQL & "        , f72m.testcd AS ITEM" & vbCr
    SQL = SQL & "        , r010m.SPCCD AS SPCCD " & vbCr
    SQL = SQL & "   FROM LJ011M j011m                                " & vbCr
    SQL = SQL & "        INNER JOIN LJ010M j010m                     " & vbCr
    SQL = SQL & "                ON j011m.bcno  = j010m.bcno         " & vbCr
    SQL = SQL & "               AND j011m.regno = j010m.regno        " & vbCr
    SQL = SQL & "        INNER JOIN LM010M r010m                     " & vbCr
    SQL = SQL & "                ON j011m.bcno   = r010m.bcno        " & vbCr
    SQL = SQL & "               AND j011m.regno  = r010m.regno       " & vbCr
    SQL = SQL & "               AND NVL(r010m.rstflg,'0') = '0'      " & vbCr
    SQL = SQL & "        INNER JOIN LF072M f72m                      " & vbCr
    SQL = SQL & "                ON f72m.eqcd    = 'G0006'           " & vbCr
    SQL = SQL & "                AND f72m.testcd  = '" & strTestCd & "' " & vbCr
    SQL = SQL & "               AND r010m.testcd = f72m.testcd       " & vbCr
    SQL = SQL & " WHERE j011m.colldt BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
    SQL = SQL & "   AND r010m.wkno between '" & pFrNo & "' AND '" & pToNo & "' " & vbCr
    SQL = SQL & "   AND j011m.spcflg  = '4'               " & vbCr
    SQL = SQL & "   AND NVL(j011m.rstflg, '0')  = '0'     " & vbCr
    SQL = SQL & "   AND j011m.bcno = '" & sBarcode & "'"
    SQL = SQL & " ORDER BY FLWKNO  " & vbCr

    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
            
            SetText .vasID, "1", .vasID.MaxRows, colCHECKBOX
            SetText .vasID, Trim(RS.Fields("접수일자")) & "", .vasID.MaxRows, colHOSPDATE
            SetText .vasID, Trim(RS.Fields("바코드번호")) & "", .vasID.MaxRows, colBARCODE
            SetText .vasID, Trim(RS.Fields("차트번호")) & "", .vasID.MaxRows, colCHARTNO
            SetText .vasID, Trim(RS.Fields("내원번호")) & "", .vasID.MaxRows, colPID
            SetText .vasID, Trim(RS.Fields("이름")) & "", .vasID.MaxRows, colPNAME
            SetText .vasID, Trim(RS.Fields("성별")) & "", .vasID.MaxRows, colPSEX
            SetText .vasID, Trim(RS.Fields("나이")) & "", .vasID.MaxRows, colPAGE
            SetText .vasID, Trim(RS.Fields("SPCCD")) & "", .vasID.MaxRows, colDISKNO
            
            '-- 화면에 표시
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = asRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    Exit For
                End If
            Next
    
            RS.MoveNext
        Loop
    
        GetSampleInfoW_JBUNIV = 1
    
    End With
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfoW_SUNGMO(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    
    
    GetSampleInfoW_SUNGMO = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
          SQL = "SELECT a.BSDATE AS BSDATE, a.SAMPLE  AS SAMPLE, a.HOSPNO AS HOSPNO" & vbCr
    SQL = SQL & "  FROM MEGADB1.TL_LABOORDER a" & vbCr
    SQL = SQL & " WHERE a.SAMPLE = '" & sBarcode & "'" & vbCr
    SQL = SQL & " GROUP BY a.SAMPLE,a.HOSPNO,a.BSDATE,a.INOUT  "

    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            SetText .vasID, "1", .vasID.MaxRows, colCHECKBOX
            'SetText .vasID, Format(Now, "yyyymmddhhmmss"), .vasID.MaxRows, colEXAMDATE
            'SetText .vasID, getMaxTestNum(Format(.dtpToday, "yyyymmdd")), .vasID.MaxRows, colSAVESEQ
            SetText .vasID, Trim(RS.Fields("BSDATE")) & "", .vasID.MaxRows, colHOSPDATE
            SetText .vasID, Trim(RS.Fields("SAMPLE")) & "", .vasID.MaxRows, colBARCODE
            SetText .vasID, Trim(RS.Fields("HOSPNO")) & "", .vasID.MaxRows, colCHARTNO
            'SetText .vasID, Trim(RS.Fields("NAME")) & "", .vasID.MaxRows, colPNAME
            'SetText .vasID, Trim(RS.Fields("SEX")) & "", .vasID.MaxRows, colPSEX
            RS.MoveNext
        Loop
    
        GetSampleInfoW_SUNGMO = 1
    
    End With
    
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfoW_MEDINOL(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    
    
    GetSampleInfoW_MEDINOL = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    If Len(sBarcode) <> 11 Then
        Exit Function
    End If
    
    If Not IsNumeric(sBarcode) Then
        Exit Function
    End If
    
'          SQL = "SELECT a.BSDATE AS BSDATE, a.SAMPLE  AS SAMPLE, a.HOSPNO AS HOSPNO" & vbCr
'    SQL = SQL & "  FROM MEGADB1.TL_LABOORDER a" & vbCr
'    SQL = SQL & " WHERE a.SAMPLE = '" & sBarcode & "'" & vbCr
'    SQL = SQL & " GROUP BY a.SAMPLE,a.HOSPNO,a.BSDATE,a.INOUT  "


          SQL = "SELECT * "
    SQL = SQL & "  FROM " & gDB_Parm.OrdTable        'MEID.VW_CPL_INTERFACE_ORDER_HO_DONG
    SQL = SQL & " WHERE specimen_ser = '" & sBarcode & "'"

    Call SetSQLData("바코드조회", SQL)
    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            SetText .vasID, "1", asRow, colCHECKBOX
            'SetText .vasID, Trim(RS.Fields("BSDATE")) & "", .vasID.MaxRows, colHOSPDATE     '처방일시
            
            SetText .vasID, Format(Now, "yyyymmdd"), asRow, colHOSPDATE    '처방일시
            
            SetText .vasID, Trim(RS.Fields("specimen_ser")) & "", asRow, colBARCODE      '바코드 번호
            SetText .vasID, Trim(RS.Fields("BUNHO")) & "", asRow, colCHARTNO      '챠트번호
            SetText .vasID, Trim(RS.Fields("SUNAME")) & "", asRow, colPNAME
            RS.MoveNext
        Loop
    
        GetSampleInfoW_MEDINOL = 1
    
    End With
    
    frmInterface.vasID.RowHeight(-1) = 12
    DoEvents
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfoW_SLALAB(ByVal asRow As Long) As Integer
    Dim sBarcode            As String
    Dim strGubun            As String
    Dim intCol              As Integer
    Dim GetOrderExamCode    As String
    Dim RS1                 As ADODB.Recordset
    Dim strRegDate          As String
    Dim lngRegNo            As Long
    
    
    GetSampleInfoW_SLALAB = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    strRegDate = "20" & Format(Mid(sBarcode, 1, 6), "##-##-##")
    lngRegNo = Val(Mid(sBarcode, 7))
    
    
    
    'sBarcode = "16040752626"
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'    If InStr(sBarcode, "-") <= 0 Then
'        Exit Function
'    End If
    
    
    '-- 바코드 번호로 오더 조회
    Dim prm1 As New ADODB.Parameter
    
    Set cmdSQL = New ADODB.Command
    Set cmdSQL.ActiveConnection = cn_Ser
    
    cmdSQL.CommandTimeout = 15
    cmdSQL.CommandText = "PG_SLA_INTERFACEMGT.SP_SLA_INTERFACEEQP_S01"
    cmdSQL.CommandType = adCmdStoredProc
    
    Set prm1 = cmdSQL.CreateParameter("in_spcno", adVarChar, adParamInput, 11, sBarcode)
    cmdSQL.Parameters.Append prm1
    
    Set RS = New ADODB.Recordset
    RS.Open cmdSQL.Execute
    
    With frmInterface
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("exam_cd")) & "',"
                
                SetText .vasID, "1", .vasID.MaxRows, colCHECKBOX
                SetText .vasID, Format(Trim(RS.Fields("bld_col_date")) & "", "yyyymmdd"), .vasID.MaxRows, colHOSPDATE
                SetText .vasID, sBarcode, .vasID.MaxRows, colBARCODE
                SetText .vasID, Trim(RS.Fields("acpno_1")) & "", .vasID.MaxRows, colCHARTNO
                SetText .vasID, Trim(RS.Fields("pt_no")) & "", .vasID.MaxRows, colPID
                SetText .vasID, Trim(RS.Fields("pt_name")) & "", .vasID.MaxRows, colPNAME
                SetText .vasID, Trim(RS.Fields("sex")) & "", .vasID.MaxRows, colPSEX
                SetText .vasID, Trim(RS.Fields("age")) & "", .vasID.MaxRows, colPAGE
                'SetText .vasID, Trim(RS.Fields("SPCCD")) & "", .vasID.MaxRows, colDISKNO
                
                '-- 화면에 표시
                For intCol = colState + 1 To .vasID.MaxCols
                    If Trim(RS.Fields("exam_cd")) = gArrEquip(intCol - colState, 3) Then
                        .vasID.Row = asRow
                        .vasID.Col = intCol
                        .vasID.BackColor = vbYellow
                        Exit For
                    End If
                Next
        
                RS.MoveNext
            Loop
        
            GetSampleInfoW_SLALAB = 1
        
        End If
    End With
        
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
    
    frmInterface.vasID.RowHeight(-1) = 12
    
    Set RS = Nothing
    Set cmdSQL = Nothing
    
End Function

'접수 - master

'    SQL = SQL & " ( :WNIFDPCD, :WNIFHPNO, :WNIFDATE, :WNIFSLIP, :WNIFITEM               " & vbCr
'    SQL = SQL & " , :WNIFOITP, :WNIFWKNO, :WNIFIDNO, :WNIFNAME, :WNIFRRNF               " & vbCr
'    SQL = SQL & " , :WNIFRSEX, :WNIFRRNS, :WNIFWARD, :WNIFROOM, :WNIFDEPT               " & vbCr
'    SQL = SQL & " , :WNIFDOCT, :WNIFMDDT, :WNIFACDT, :WNIFACTM, :WNIFRQDT               " & vbCr
'    SQL = SQL & " , :WNIFRQTM, :WNIFLSDT, :WNIFRPDT, :WNIFRPTM, :WNIFRPMT               " & vbCr
'    SQL = SQL & " , :WNIFRPMD, :WNIFSMDT, :WNIFSMTM, :WNIFSMPL, :WNIFSMGB               " & vbCr
'    SQL = SQL & " , :WNIFSMYR, :WNIFSMSN, :WNIFSMS1, :WNIFSMS2, :WNIFSTAT               " & vbCr
'    SQL = SQL & " , :WNIFQCHK, :WNIFSPCL, :WNIFCONT, :WNIFODNU, :WNIFTYPE               " & vbCr
'    SQL = SQL & " , :WNIFJSNU, :WNIFJSCD, :WNIFDOID, :WNIFDONM, :WNIFTRFG               " & vbCr
'    SQL = SQL & " , :WNIFPRDT, :WNIFPRTM, :WNIFMEMO, :WNIFECT1, :WNIFECT2               " & vbCr
'    SQL = SQL & " , :WNIFECT3, :WNIFCHRT, :WNIFSLID, :WNIFBARC, :WNIFACUS )"

'    cmdSQL.Parameters.Append cmdSQL.CreateParameter("WNIFDPCD", adVarChar, , 20, "LA")
'        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
'                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)

Function Reg_acawnifh_ILSIN() As Boolean
                
    
On Error GoTo ErrHandle

    Reg_acawnifh_ILSIN = False
    
    mOCS.FWorkNo = GetNextWkno
    mOCS.FBarCode = Format(Now, "yy") & Mid(GetNextSpcid, 1, 6) + "11"

    '-  wnifacdt/tm(접수일자/시간)
    '-  wnifrqdt/tm(의뢰일자/시간)
    '-  wnifrpdt/tm(보고일자/시간)
    
          SQL = "Insert Into  acawnifh                                                  " & vbCrLf
    SQL = SQL & " ( WNIFDPCD, WNIFHPNO, WNIFDATE, WNIFSLIP, WNIFITEM                    " & vbCrLf      '5
    SQL = SQL & " , WNIFOITP, WNIFWKNO, WNIFIDNO, WNIFNAME, WNIFRRNF                    " & vbCrLf      '10
    SQL = SQL & " , WNIFRSEX, WNIFRRNS, WNIFWARD, WNIFROOM, WNIFDEPT                    " & vbCrLf      '15
    SQL = SQL & " , WNIFDOCT, WNIFMDDT, WNIFACDT, WNIFACTM, WNIFRQDT                    " & vbCrLf      '16/17/18(접수일자/시간)19/20(의뢰일자/시간)
    SQL = SQL & " , WNIFRQTM, WNIFLSDT, WNIFRPDT, WNIFRPTM, WNIFRPMT                    " & vbCrLf      '22: 보고일자/24: 보고시간/
    SQL = SQL & " , WNIFRPMD, WNIFSMDT, WNIFSMTM, WNIFSMPL, WNIFSMGB                    " & vbCrLf      '30
    SQL = SQL & " , WNIFSMYR, WNIFSMSN, WNIFSMS1, WNIFSMS2, WNIFSTAT                    " & vbCrLf      '35 ipslsmsn (7자리)
    SQL = SQL & " , WNIFQCHK, WNIFSPCL, WNIFCONT, WNIFODNU, WNIFTYPE                    " & vbCrLf      '40
    SQL = SQL & " , WNIFJSNU, WNIFJSCD, WNIFDOID, WNIFDONM, WNIFTRFG                    " & vbCrLf      '45
    SQL = SQL & " , WNIFPRDT, WNIFPRTM, WNIFMEMO, WNIFECT1, WNIFECT2                    " & vbCrLf      '50
    SQL = SQL & " , WNIFECT3, WNIFCHRT, WNIFSLID, WNIFBARC, WNIFACUS )                  " & vbCrLf      '55
    SQL = SQL & " Values (                                                              " & vbCrLf
    
    SQL = SQL & "   'LA'                                " & vbCrLf
    SQL = SQL & " , '01'                                " & vbCrLf
    SQL = SQL & " , '" & Format(Now, "yyyymmdd") & "'   " & vbCrLf
    SQL = SQL & " , 'LAE'                               " & vbCrLf
    SQL = SQL & " , '00'                                " & vbCrLf    '5
    
    SQL = SQL & " , 'A'                                 " & vbCrLf
    SQL = SQL & " , '" & mOCS.FWorkNo & "'              " & vbCrLf
    SQL = SQL & " , '" & mOCS.FPID & "'                 " & vbCrLf
    SQL = SQL & " , '" & mOCS.FPNM & "'                 " & vbCrLf
    SQL = SQL & " , '" & mOCS.FJNO & "'                 " & vbCrLf   '10
    
    SQL = SQL & " , '" & mOCS.FJNO1 & "'                " & vbCrLf
    SQL = SQL & " , '" & mOCS.FJNO2 & "'                " & vbCrLf
    SQL = SQL & " , '" & mOCS.FWard & "'                " & vbCrLf
    SQL = SQL & " , '" & mOCS.FRoom & "'                " & vbCrLf
    SQL = SQL & " , '" & mOCS.FDept & "'                " & vbCrLf  '15
    
    SQL = SQL & " , '" & mOCS.FDoct & "'                " & vbCrLf
    SQL = SQL & " , '" & mOCS.FMdDt & "'                " & vbCrLf  '17
    SQL = SQL & " , '" & mOCS.FAcDt & "'                " & vbCrLf  '18     접수일자
    SQL = SQL & " , '" & mOCS.FAcTm & "'                " & vbCrLf  '19     접수시간
    SQL = SQL & " , '" & mOCS.FRqDt & "'                " & vbCrLf  '20     의뢰일자
    
    SQL = SQL & " , '" & mOCS.FRqTm & "'                " & vbCrLf  '21     의뢰시간
    SQL = SQL & " , '" & Format(Now, "yyyymmdd") & "'   " & vbCrLf
    SQL = SQL & " , '" & Format(Now, "yyyymmdd") & "'   " & vbCrLf  '23     보고일자
    SQL = SQL & " , '" & Format(Now, "hhmm") & "'       " & vbCrLf  '24     보고시간
    SQL = SQL & " , '" & mOCS.FDocID & "'               " & vbCrLf  '25
    
    SQL = SQL & " , '" & mOCS.FDocID & "'               " & vbCrLf  '
    SQL = SQL & " , '" & Format(Now, "yyyymmdd") & "'   " & vbCrLf
    SQL = SQL & " , '" & Format(Now, "hhmm") & "'       " & vbCrLf
    SQL = SQL & " , '004'                               " & vbCrLf
    SQL = SQL & " , ''                                  " & vbCrLf  '30
    
    SQL = SQL & " , '" & Format(Now, "yy") & "'         " & vbCrLf
    SQL = SQL & " , '" & Mid(mOCS.FBarCode, 3, 7) & "'  " & vbCrLf
    SQL = SQL & " , 1                                   " & vbCrLf
    SQL = SQL & " , 1                                   " & vbCrLf
    SQL = SQL & " , '2'                                 " & vbCrLf  '35
    
    SQL = SQL & " , '-'                                 " & vbCrLf
    SQL = SQL & " , '-'                                 " & vbCrLf
    SQL = SQL & " , '-'                                 " & vbCrLf
    SQL = SQL & " , 15                                  " & vbCrLf
    SQL = SQL & " , 'I'                                 " & vbCrLf  '40

    SQL = SQL & " , 15                                  " & vbCrLf
    SQL = SQL & " , '" & gAllExamCD & "'                " & vbCrLf
    SQL = SQL & " , ''                                  " & vbCrLf
    SQL = SQL & " , ''                                  " & vbCrLf
    SQL = SQL & " , 'Y'                                 " & vbCrLf  '45
    
    SQL = SQL & " , ''                                  " & vbCrLf
    SQL = SQL & " , ''                                  " & vbCrLf
    SQL = SQL & " , ''                                  " & vbCrLf
    SQL = SQL & " , ''                                  " & vbCrLf
    SQL = SQL & " , ''                                  " & vbCrLf  '50
    
    SQL = SQL & " , ''                                  " & vbCrLf
    SQL = SQL & " , 'LA'                                " & vbCrLf
    SQL = SQL & " , 'NICU'                              " & vbCrLf
    SQL = SQL & " , ''                                  " & vbCrLf
    SQL = SQL & " , '" & mOCS.FDocID & "')              " & vbCrLf  '55
    
    Call SetSQLData("3_접수저장", SQL)
    Call SetRawData("3_접수저장 : " & SQL & vbCrLf)

    Res = SendQuery(gServer, SQL)
    
    If Res > 0 Then
        Reg_acawnifh_ILSIN = True
    End If
Exit Function

ErrHandle:
    
    MsgBox "저장오류", vbOKOnly, "서버저장"
    
End Function
    
'-- 검사자 정보 가져오기
Function GetSampleInfoW_ILSIN(ByVal asRow As Long) As Integer
    Dim sBarcode            As String
    Dim strGubun            As String
    Dim intCol              As Integer
    Dim GetOrderExamCode    As String
    Dim RS1                 As ADODB.Recordset
    Dim strRegDate          As String
    Dim lngRegNo            As Long
    
    Dim strDate             As String
    Dim strDateN1           As String
    Dim strDateN2           As String
    
    GetSampleInfoW_ILSIN = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    If sBarcode = "" Then
        Exit Function
    End If
    
    
    strDateN2 = Format(Now - 2, "yyyymmdd")
    strDateN1 = Format(Now - 1, "yyyymmdd")
    strDate = Format(Now, "yyyymmdd")
    
    '-  wnifacdt/tm(접수일자/시간) <- ipslacdt/tm 를 입력
    '-  wnifrqdt/tm(의뢰일자/시간) <- ipslgndt/tm 를 입력
    
    
    '-- 바코드 번호로 오더 조회
          SQL = "Select ptbsname                    " & vbCrLf
    SQL = SQL & "     , substr(ptbsrrns,1,1) JNO1   " & vbCrLf  '성별
    SQL = SQL & "     , substr(ptbsrrns,2,6) JNO2   " & vbCrLf  '나이
    SQL = SQL & "     , ptbsrrnf JNO                " & vbCrLf
    SQL = SQL & "     , ipslward                    " & vbCrLf
    SQL = SQL & "     , ipslroom                    " & vbCrLf
    SQL = SQL & "     , ipsldept                    " & vbCrLf
    SQL = SQL & "     , ipsldoct                    " & vbCrLf
    SQL = SQL & "     , ipslgnus                    " & vbCrLf      '의사사번 - 보고자.
    SQL = SQL & "     , ipslmddt                    " & vbCrLf
    SQL = SQL & "     , ipslacdt                    " & vbCrLf      '접수일자
    SQL = SQL & "     , ipslactm                    " & vbCrLf      '접수시간
    SQL = SQL & "     , ipslhpdt                    " & vbCrLf
    SQL = SQL & "     , ipslhptm                    " & vbCrLf
    SQL = SQL & "     , ipslgndt                    " & vbCrLf      '의뢰일자
    SQL = SQL & "     , ipslgntm                    " & vbCrLf      '의뢰시간
    SQL = SQL & "  From pmcptbsm, ocsipslh          " & vbCrLf
    SQL = SQL & " Where ipslidno = '" & sBarcode & "'" & vbCrLf     '챠트번호
    SQL = SQL & "   And ipslmddt IN ('" & strDateN1 & "','" & strDate & "') " & vbCrLf  '처방일자
    SQL = SQL & "   And ipslcode = '" & gOrdCd & "' " & vbCrLf '처방코드 'SUEBGAG'
    SQL = SQL & "   And ipslidno = ptbsidno         " & vbCrLf
    SQL = SQL & "   And ipslflag = 'O'              " & vbCrLf
    
    '2019-12-09
'    SQL = SQL & "   And ipslstat = '0'              " & vbCrLf  '7
                
    Call SetSQLData("1_환자조회", SQL)
    Call SetRawData("1_환자조회 : " & SQL & vbCrLf)

    Set RS = cn_Ser.Execute(SQL)

    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With mOCS
                .FPID = sBarcode
                .FOrdCd = gOrdCd
                .FPNM = Trim(RS.Fields("ptbsname")) & ""
                .FJNO = Trim(RS.Fields("JNO")) & ""
                .FJNO1 = Trim(RS.Fields("JNO1")) & ""
                .FJNO2 = Trim(RS.Fields("JNO2")) & ""
                .FWard = Trim(RS.Fields("ipslward")) & ""
                .FRoom = Trim(RS.Fields("ipslroom")) & ""
                .FDept = Trim(RS.Fields("ipsldept")) & ""
                .FMdDt = Trim(RS.Fields("ipslmddt")) & ""
                .FAcDt = Trim(RS.Fields("ipslacdt")) & ""   '접수일자   acawnifh 테이블 ipslacdt(접수일자) 에 업데이트
                .FAcTm = Trim(RS.Fields("ipslactm")) & ""   '접수시간   acawnifh 테이블 ipslactm(접수시간) 에 업데이트
                .FRqDt = Trim(RS.Fields("ipslgndt")) & ""   '의뢰일자   acawnifh 테이블 ipslgndt(의뢰일자) 에 업데이트
                .FRqTm = Trim(RS.Fields("ipslgntm")) & ""   '의뢰시간   acawnifh 테이블 ipslgntm(의뢰시간) 에 업데이트
                .FHpDt = Trim(RS.Fields("ipslhpdt")) & ""
                .FHpTm = Trim(RS.Fields("ipslhptm")) & ""
                .FDoct = Trim(RS.Fields("ipsldoct")) & ""
                .FDocID = Trim(RS.Fields("ipslgnus")) & ""
            End With
            
            RS.MoveNext
        Loop
    
        GetSampleInfoW_ILSIN = 1
    
    End If
            
    frmInterface.vasID.RowHeight(-1) = 12
    
    Set RS = Nothing
    
    With mOCS
        SQL = sBarcode & vbCr
        SQL = SQL & .FOrdCd & vbCr
        SQL = SQL & .FPNM & vbCr
        SQL = SQL & .FJNO & vbCr
        SQL = SQL & .FJNO1 & vbCr
        SQL = SQL & .FJNO2 & vbCr
        SQL = SQL & .FWard & vbCr
        SQL = SQL & .FRoom & vbCr
        SQL = SQL & .FDept & vbCr
        SQL = SQL & .FMdDt & vbCr
        SQL = SQL & .FAcDt & vbCr
        SQL = SQL & .FAcTm & vbCr
        SQL = SQL & .FHpDt & vbCr
        SQL = SQL & .FHpTm & vbCr
        SQL = SQL & .FDoct & vbCr
        SQL = SQL & .FRqDt & vbCr
        SQL = SQL & .FRqTm & vbCr
        SQL = SQL & .FDocID & vbCr
    End With
    
10:47:55 2_환자조회결과 :
'06870254
'SUEBGAG
'한가희애기
'200405
'3
'0
'PD
'300
'PD
'20200408
'-
'-
'-
'-
'PD07
'20200408
'948
'10136


'04:57:04 2_환자조회결과 :
'06870261
'SUEBGAG
'손아름애기
'200406
'3
'0
30
300
PD
20200408




PD07
20200407
837
10136


    
    Call SetSQLData("2_환자조회결과", SQL)
    Call SetRawData("2_환자조회결과 : " & SQL & vbCrLf)

End Function

'-- 검사자 정보 가져오기
Function GetSampleInfoW_JAINCOM(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    
    
    GetSampleInfoW_JAINCOM = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'      -- 테이블 사용
          SQL = "SELECT DiSTINCT b.SCP42JDATE as 접수일자, a.SCP41SPMNO2 as 바코드번호, b.SCP42IDNOA as 내원번호, a.SCP41NAME as 이름, a.SCP41SEX as 성별, a.SCP41BIRTH as 나이,b.SCP42SUGACD as ITEM"
    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & sBarcode & "'"
    If frmInterface.chkSaveAll.Value = "0" Then
    '    SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
        SQL = SQL & vbCrLf & "   AND (b.SCP42RESULT IS NULL OR b.SCP42RESULT = '') "
    End If

    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
            
            SetText .vasID, "1", .vasID.MaxRows, colCHECKBOX
            SetText .vasID, Trim(RS.Fields("접수일자")) & "", .vasID.MaxRows, colHOSPDATE
            SetText .vasID, Trim(RS.Fields("바코드번호")) & "", .vasID.MaxRows, colBARCODE
            'SetText .vasID, Trim(RS.Fields("차트번호")) & "", .vasID.MaxRows, colCHARTNO
            SetText .vasID, Trim(RS.Fields("내원번호")) & "", .vasID.MaxRows, colPID
            SetText .vasID, Trim(RS.Fields("이름")) & "", .vasID.MaxRows, colPNAME
            SetText .vasID, Trim(RS.Fields("성별")) & "", .vasID.MaxRows, colPSEX
            SetText .vasID, Trim(RS.Fields("나이")) & "", .vasID.MaxRows, colPAGE
            'SetText .vasID, Trim(RS.Fields("SPCCD")) & "", .vasID.MaxRows, colDISKNO
            
            '-- 화면에 표시
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = asRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    Exit For
                End If
            Next
    
            RS.MoveNext
        Loop
    
        GetSampleInfoW_JAINCOM = 1
    
    End With
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function
    
'-- 검사자 정보 가져오기
Function GetSampleInfoW_BIT(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    
    GetSampleInfoW_BIT = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
     
          'SQL = "SELECT DISTINCT P.PbsPatNam, O.OSPCHTNUM, E.LabShtNam, R.ResOcmNum, R.ResLabCod, R.ResOdrSeq, R.ResSeq, R.ResSubSeq, R.ResRltVal" & vbCrLf
          SQL = "SELECT DISTINCT P.PbsPatNam, O.OSPCHTNUM, R.ResOcmNum, R.ResLabCod AS EXAMCODE , R.ResOdrSeq, R.ResSeq, R.ResSubSeq " & vbCrLf
    SQL = SQL & "  FROM RsbInf M, ResInf R, ospinf O, PBSINF P, LabMst E" & vbCrLf
    SQL = SQL & " WHERE M.RsbBarCod = '" & sBarcode & "'" & vbCrLf
    SQL = SQL & "   And M.RsbAckStt <> 'A' " & vbCrLf
    SQL = SQL & "   And O.OspChkStt <> 'F' " & vbCrLf
    SQL = SQL & "   And (R.ResRepTyp Is Null or R.ResRepTyp <> 'F') " & vbCrLf
    SQL = SQL & "   And (R.ResRltVal is Null or R.ResRltVal = '') " & vbCrLf
    SQL = SQL & "   And R.ResStatus <> '5' " & vbCrLf
    SQL = SQL & "   And M.RSBACPNUM = R.ResRsbAcp" & vbCrLf
    SQL = SQL & "   And R.ResOcmNum = O.OspOcmNum" & vbCrLf
    SQL = SQL & "   and R.ResOdrSeq = O.OspOdrSeq" & vbCrLf
    SQL = SQL & "   and R.ResSeq    = O.OspSeq" & vbCrLf
    SQL = SQL & "   and O.OSPCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   and R.ResLabcod = E.LabCod" & vbCrLf
    
    SetRawData "[SQL1]" & SQL
    
    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("EXAMCODE")) & "',"
            
            SetText .vasID, "1", asRow, colCHECKBOX
            SetText .vasID, sBarcode, asRow, colBARCODE
            SetText .vasID, Trim(RS.Fields("OSPCHTNUM")), asRow, colCHARTNO         '챠트번호(결과상태 저장시 필요)
            SetText .vasID, Trim(RS.Fields("ResOcmNum")), asRow, colPID             '등록번호(결과     저장시 필요)
            SetText .vasID, Trim(RS.Fields("PbsPatNam")), asRow, colPNAME           '환자명
            
            
            'SetText .vasID, "12345", asRow, colCHARTNO         '챠트번호
            'SetText .vasID, "67890", asRow, colPID            '등록번호(저장시 필요)
            'SetText .vasID, "홍길릴", asRow, colPNAME           '환자명
            
            '-- 화면에 표시
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = asRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    '-- 결과저장용 SEQ
                    gArrEquip(intCol - colState, 7) = Trim(RS.Fields("ResOdrSeq")) & "|" & Trim(RS.Fields("ResSeq")) & "|" & Trim(RS.Fields("ResSubSeq"))   '결과저장용 번호's
                    'gArrEquip(intCol - colState, 7) = "987" & "|" & "654" & "|" & "321"
                    Exit For
                End If
            Next
    
            RS.MoveNext
        Loop
    
        GetSampleInfoW_BIT = 1
    
    End With
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    frmInterface.vasID.RowHeight(-1) = 12

End Function

Public Function GetSameRowNum(ByVal strBarNo As String) As Integer
    Dim i As Integer

    GetSameRowNum = 0
    With frmInterface.vasID
        For i = 1 To .MaxRows
            .Row = i
            .Col = colBARCODE
            If Trim(.Text) = strBarNo Then
                GetSameRowNum = i
                Exit Function
            End If
        Next
    End With
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfoW_GINUSDLL(ByVal asRow As Long) As Integer
    Dim pBarNo  As String
    Dim i       As Integer
    Dim intCol  As Integer
    Dim strItem As String
    
    '-- 지누스
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    GetSampleInfoW_GINUSDLL = -1
    
    pBarNo = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If pBarNo = "" Then
        Exit Function
    End If
    
    '-- 검사ITEM 가져오기
                 strRequest = "jobs" + vbTab + "Q" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "smp_no" + vbTab + pBarNo + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + gGINUS_Parm.MCD + vbTab + vbCr
    
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- 바코드로 검사대상 조회(https://211.172.17.66)
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    With frmInterface.vasID
        If UBound(varResponse) > 0 Then
            For i = 0 To UBound(varResponse) - 1
                SetText frmInterface.vasID, "1", asRow, colCHECKBOX
                SetText frmInterface.vasID, Mid(mGetP(varResponse(i), 25, vbTab), 1, 8), asRow, colHOSPDATE
                SetText frmInterface.vasID, mGetP(varResponse(i), 0, vbTab), asRow, colBARCODE
                SetText frmInterface.vasID, mGetP(varResponse(i), 7, vbTab), asRow, colPID
                SetText frmInterface.vasID, mGetP(varResponse(i), 26, vbTab), asRow, colPNAME
                
                Select Case mGetP(varResponse(i), 29, vbTab)
                    Case "O": SetText frmInterface.vasID, "외래", asRow, colINOUT
                    Case "E": SetText frmInterface.vasID, "응급", asRow, colINOUT
                    Case "I": SetText frmInterface.vasID, "입원", asRow, colINOUT
                End Select
                
                
                For intCol = colState + 1 To .MaxCols
                    If mGetP(varResponse(i), 6, vbTab) = gArrEquip(intCol - colState, 3) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        '-- 결과저장용 SEQ
                        gArrEquip(intCol - colState, 7) = mGetP(varResponse(i), 3, vbTab) & "|" & mGetP(varResponse(i), 4, vbTab) & "|" & mGetP(varResponse(i), 5, vbTab)
                        Exit For
                    End If
                Next intCol
            Next i
        Else
            SetText frmInterface.vasID, "No Order", asRow, colState
        End If
    End With
    
    GetSampleInfoW_GINUSDLL = 1

'          SQL = " SELECT DISTINCT REQ_DT AS 접수일자"
'    SQL = SQL & ", LOT_NO AS 차트번호"
'    SQL = SQL & ", REQ_SEQ AS 내원번호"
'    SQL = SQL & ", '입원' AS 입외"
'    SQL = SQL & ", '홍길동' AS 이름"
'    SQL = SQL & ", '남자' AS 성별"
'    SQL = SQL & ", REQ_SEQ AS 나이" & vbCrLf
'    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
'    SQL = SQL & " WHERE QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
'    SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
'
'    Res = GetDBSelectColumn(gServer, SQL)
        
'    If Res = 1 Then
'        SetText frmInterface.vasID, "1", asRow, colCheckBox
'        SetText frmInterface.vasID, sBarcode, asRow, colBARCODE
'        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colHOSPDATE
'        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colCHARTNO
'        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID
'        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colINOUT
'        SetText frmInterface.vasID, Trim(gReadBuf(4)), asRow, colPNAME
'        SetText frmInterface.vasID, Trim(gReadBuf(5)), asRow, colPSEX
'        SetText frmInterface.vasID, Trim(gReadBuf(6)), asRow, colPAGE
'
'        GetSampleInfoW_GINUSDLL = 1
'
'    Else
'        GetSampleInfoW_GINUSDLL = -1
'    End If
'
'    frmInterface.vasID.RowHeight(-1) = 12

End Function

'Function GetSampleInfoR(ByVal asRow As Long) As Integer
'    Dim sBarcode As String
'    Dim sSpecNo As String
'
'    GetSampleInfoR = -1
'
'    '-- 환자정보 가져오기
'    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBARCODE))   '샘플 바코드 번호
'
'    If sBarcode = "" Then
'        Exit Function
'    End If
'
'    '-- 바코드번호로 환자정보 불러오기
'          SQL = "SELECT " & gDBCOLUMN_Parm.PID & "," & gDBCOLUMN_Parm.PNAME & "," & gDBCOLUMN_Parm.PSEX & "," & gDBCOLUMN_Parm.PAGE & vbCrLf
'    SQL = SQL & "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
'    SQL = SQL & " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
'    If gDBCOLUMN_Parm.STATUS <> "" Then
'        SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
'    End If
'    If gDBCOLUMN_Parm.RESULT <> "" Then
'        SQL = SQL + "   AND (" & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL)"
'    End If
'
'    Res = GetDBSelectColumn(gServer, SQL)
'
'    If Res = 1 Then
'        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo
''        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
'        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPNAME
'        '-- 성별이 없을경우 주민번호로 찾기
'        'strSex = IIf(Mid(Trim(gReadBuf(4)), 7, 1) = "1", "M", "F")
'        'SetText frmInterface.vasID, strSex, colSex    '7  성별
''        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '7  성별
'        '-- 나이가 없을경우 주민번호로 찾기
'        'strAge = Format(Now, "yyyy") - Mid(Trim(gReadBuf(3)), 1, 4)
'        'SetText frmInterface.vasID, strAge, asRow, colAge
''        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colSex    '8  나이
'
'        GetSampleInfoR = 1
'    Else
'
'        GetSampleInfoR = -1
'    End If
'
'End Function

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
    
    SQL = " Select examcode From EQPMASTER " & vbCrLf & _
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
    SQL = SQL & "  From EQPMASTER "
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

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Function GetOrderExamCode(argEquipCode As String, argPID As String) As String
    Dim RS As ADODB.Recordset
    Dim intCol As Integer

    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
          

     
          SQL = "SELECT P.PbsPatNam, O.OSPCHTNUM, R.ResLabCod, E.LabShtNam, R.ResOcmNum, R.ResOdrSeq, R.ResSeq, R.ResSubSeq, R.ResRltVal" & vbCrLf
    SQL = SQL & "  FROM RsbInf M, ResInf R, ospinf O, PBSINF P, LabMst E" & vbCrLf
    SQL = SQL & " WHERE M.RsbBarCod = '" & argPID & "'" & vbCrLf
    SQL = SQL & "   And M.RsbAckStt <> 'A' " & vbCrLf
    SQL = SQL & "   And O.OspChkStt <> 'F' " & vbCrLf
    SQL = SQL & "   And (R.ResRepTyp Is Null or R.ResRepTyp <> 'F') " & vbCrLf
    SQL = SQL & "   And (R.ResRltVal is Null or R.ResRltVal = '') " & vbCrLf
    SQL = SQL & "   And R.ResStatus <> '5' " & vbCrLf
    SQL = SQL & "   And M.RSBACPNUM = R.ResRsbAcp" & vbCrLf
    SQL = SQL & "   And R.ResOcmNum = O.OspOcmNum" & vbCrLf
    SQL = SQL & "   and R.ResOdrSeq = O.OspOdrSeq" & vbCrLf
    SQL = SQL & "   and R.ResSeq    = O.OspSeq" & vbCrLf
    SQL = SQL & "   and O.OSPCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   and R.ResLabcod = E.LabCod" & vbCrLf
    
    Set RS = cn_Ser.Execute(SQL)
    
    Do Until RS.EOF
        GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("EXAMCODE")) & "',"
        
        '-- 화면에 표시
        With frmInterface
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = .vasID.ActiveRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    Exit For
                End If
            Next
        End With
        
        RS.MoveNext
    Loop
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
    End If
    
    RS.Close
    
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
    
    argPID = Mid(argPID, 1, 10)
    
          SQL = "SELECT DiSTINCT b.SCP42SUGACD "
    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & argPID & "'"
    SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
    
    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_New = GetOrderExamCode_New & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
    If GetOrderExamCode_New <> "" Then
        GetOrderExamCode_New = Mid(GetOrderExamCode_New, 1, Len(GetOrderExamCode_New) - 1)
    End If
    
End Function

Function GetOrderExamCode_UBCARE(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String
Dim rs_svr As ADODB.Recordset

    GetOrderExamCode_UBCARE = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If

          SQL = "SELECT DiSTINCT EXAMCODE "
    SQL = SQL & vbCrLf & "  FROM PATRESULT "
    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & argPID & "'"
'    SQL = SQL & vbCrLf & "   AND (SAVESEQ IS NULL OR SAVESEQ = '')"
    
    Set rs_svr = cn.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_UBCARE = GetOrderExamCode_UBCARE & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
    If GetOrderExamCode_UBCARE <> "" Then
        GetOrderExamCode_UBCARE = Mid(GetOrderExamCode_UBCARE, 1, Len(GetOrderExamCode_UBCARE) - 1)
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
    SQL = SQL & "  From EQPMASTER "
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
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 샘플 바코드 번호
    
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
    SQL = SQL & "  FROM EQPMASTER "
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

'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetGetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_AU480 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    
    ClearSpread frmInterface.vasTemp1
    
    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            'AU480의 경우 장비에서 dilution 사용시 끝에 '0'추가
            strExamCode = strExamCode & "0" & Trim(gReadBuf(i)) & "0"
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_AU480 = strExamCode
    
End Function


'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetGetEquipExamCode_CentaurCP(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_CentaurCP = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""

    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_CentaurCP = strExamCode
    
End Function

'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetGetEquipExamCode_Hitachi7080(argEquipCode As String, argPID As String, Optional intRow As Long) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim strExamCode() As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_Hitachi7080 = ""
    j = 0
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    'strExamCode = ""

    'ReDim strExamCode(UBound(gReadBuf))
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            ReDim Preserve strExamCode(j)
            strExamCode(j) = Trim(gReadBuf(i))
            j = j + 1
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_Hitachi7080 = strExamCode
    
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
    SQL = SQL & "  From EQPMASTER "
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


