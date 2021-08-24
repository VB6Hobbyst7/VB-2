Attribute VB_Name = "modDbLibrary"
Option Explicit


Public Function db_select_Vas(argServer As Integer, argSQL As String, ByVal argSpread As vaSpread, Optional argRow As Integer = 1, Optional argCol As Integer = 1) As Integer
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
  
    If argSpread.MaxCols < RS.Fields.Count + argCol - 1 Then
        argSpread.MaxCols = RS.Fields.Count + argCol - 1
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
            argSpread.Col = j + argCol
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
    Dim strAllResult  As String
    Dim sParam As String
    Dim sRcvData As String
    Dim sCommt  As String
    Dim sComment  As String
    
    sComment = ""
    sComment = sComment & "■ 1＋  :  100 copy/reaction 이하%0D"
    sComment = sComment & "   2＋  :  100 copy/reaction ~ 100,000 copy/reaction%0D"
    sComment = sComment & "   3＋  :  100,000 copy/reaction 이상%0D%0D"
    sComment = sComment & "■ 검사자  :  김지영   ■ 판독의  :  허미나.M.D%0D"
    sComment = sComment & "위 검사결과는 진단검사의학과 전문의가 판독, 검증하였습니다."

'On Error GoTo ErrHandle
On Error Resume Next

    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strChtNum = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        strExamDate = Trim(GetText(.vasID, argSpcRow, colEXAMDATE))
        strSaveSeq = Trim(GetText(.vasID, argSpcRow, colSAVESEQ))
        
        '-- Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        
'              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE,INOUT,EXAMNAME " & vbCrLf
              SQL = "SELECT DISTINCT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE,INOUT,EXAMNAME " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '장비코드
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'  " & vbCrLf                                      '검사일
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf       '바코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq       '저장번호
'        SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colDISKNO)) & "' " & vbCrLf         'DISK 번호(진료검사ID)
'        SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPOSNO)) & "' "                    'POS 번호(진료지원ID)
        SQL = SQL & " ORDER BY SEQNO"
        
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
        
'        strInCD = ""
'        strInVal = ""
'        intTotCnt = 0
        
'        cn_Ser.BeginTrans
        strAllResult = ""
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            'strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            strEqpCd = Trim(GetText(.vasTemp, iRow, 1))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))     '결과(IU/ml)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))     '결과(Class)
            'strRefVal = Trim(GetText(.vasTemp, iRow, 5))    '판정
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            'http://kis.kuh.ac.kr:80/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetIFGyeolGwaDR&Data1=12943007070L2111516320130122113646&Data2=J04
            'http://kis.kuh.ac.kr/   lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetIFGyeolGwaDR&Data1=32942000940VB8506AN20160503112333&Data2=CFX96


'GTData1=14127002830HB2CAP96|< 120 Copies/mL (< 20 IU/mL)|<|00000000000000&Data2=P01
'GTData1=14127003600HB2CAP96|4.71E+02 Copies/mL (8.1E+01 IU/mL)||00000000000000&Data2=P01
'GTData1=14127003820HB2CAP96|2.58E+04 Copies/mL (4.43E+03 IU/mL)||00000000000000&Data2=P01GTData1=14127005630HB2CAP96|< 120 Copies/mL (< 20 IU/mL)|<|00000000000000&Data2=P01GTData1=14127006050HB2CAP96|> 9.89E+08 Copies/mL (> 1.7E+08 IU/mL)|>|00000000000000&Data2=P01GTData1=14127008660HB2CAP96|> 9.89E+08 Copies/mL (> 1.7E+08 IU/mL)|>|00000000000000&Data2=P01GTData1=14127015280HB2CAP96|< 120 Copies/mL (< 20 IU/mL)|<|00000000000000&Data2=P01GTData1=14127022460HB2CAP96|< 120 Copies/mL (< 20 IU/mL)|<|00000000000000&Data2=P01GTData1=14128004220HB2CAP96|Target Not Detected||00000000000000&Data2=P01GTData1=14128003980HB2CAP96|< 120 Copies/mL (< 20 IU/mL)|<|00000000000000&Data2=P01

            If InStr(sResult, "＋＋＋") > 0 Then
                sCommt = "100,000 copy/reaction 이상"
            ElseIf InStr(sResult, "＋＋") > 0 Then
                sCommt = "100 copy/reaction ~ 100,000 copy/reaction 이하"
            ElseIf InStr(sResult, "＋") > 0 Then
                sCommt = "100 copy/reaction 이하"
            Else
                sCommt = ""
            End If
            
            '-- 일반검체
'            sParam = "http://" & gKNUH.IP & gKNUH.RSTURL & "&"
'            sParam = sParam & "Data1=" & lsID
'            sParam = sParam & "" & strEqpCd
'            sParam = sParam & "" & sResult
'            sParam = sParam & "" & sCommt
'            sParam = sParam & "" & Format(Now, "yyyymmddhhmmss")
'            sParam = sParam & "&Data2=" & gEquipCode
        
'기존에 사용하시던 URL뒤에 바코드를 제외한 정보들을 붙이시면 됩니다.
'
'바코드번호(%03)
'
'+검사코드(%17)+결과값(%17)+IF코멘트(%17)+검사일시(%17)+결과구분(%17)+검사코멘트(%03)
'+검사코드(%17)+결과값(%17)+IF코멘트(%17)+검사일시(%17)+결과구분(%17)+검사코멘트(%03)
'+검사코드(%17)+결과값(%17)+IF코멘트(%17)+검사일시(%17)+결과구분(%17)+검사코멘트(%03)
'+검사코드(%17)+결과값(%17)+IF코멘트(%17)+검사일시(%17)+결과구분(%17)+검사코멘트(%03)
'+검사코드(%17)+결과값(%17)+IF코멘트(%17)+검사일시(%17)+결과구분(%17)+검사코멘트(%03)
'+검사코드(%17)+결과값(%17)+IF코멘트(%17)+검사일시(%17)+결과구분(%17)+검사코멘트(%03)
'
'현재 사용중인 다른 검사입니다.(소견없는 검사입니다.)
'
'12762012350 %03a-IgE%1770.8%17%1720120724132008%03
'
'd1%170.04%20(0)%17%1720120724132008%03d2%170.01%20(0)%17%1720120724132008%03
'f1%170.18%20(0)%17%1720120724132008%03f2%170.27%20(0)%17%1720120724132008%03
'f4%170.04%20(0)%17%1720120724132008%03f11%170.02%20(0)%17%1720120724132008%03
'f13%170.01%20(0)%17%1720120724132008%03
'f14%170.02%20(0)%17%1720120724132008%03 &Data2=S05
'
'지난번에 주신 검사로 만든 정보입니다.
'
'Data1=
'
'14136012640%03
'
'NG%17Positive(2＋)%17%1720160526145625%17100 copy reatn 이하%03
'
'TV%17Positive(2＋)%17%1720160526145625%17%03&Data2=CF2
        
        
            '-- 일반검체(코멘트적용)
            sParam = "http://" & gKNUH.IP & gKNUH.RSTURL & "&"
            sParam = sParam & "Data1=" & lsID
            sParam = sParam & "" & strEqpCd
            sParam = sParam & "" & sResult
            sParam = sParam & "" & "" 'sCommt
            sParam = sParam & "" & Format(Now, "yyyymmddhhmmss")
            sParam = sParam & "" & ""
            If iRow = 1 Then
                sParam = sParam & "" & sComment
            Else
                sParam = sParam & ""
            End If
            sParam = sParam & "&Data2=" & gEquipCode
            
            
            '==> 서버로 결과저장
            'Print #1, vbNewLine & "[sUrl]" & sUrl;
            SetRawData "[sParam]" & sParam
            
            sRcvData = OpenURLWithIE2(sParam, frmInterface.Inet1)
            'sRcvData = Inet1.OpenURL(sUrl & sParam)
    
            SetRawData "[IN]" & sRcvData
                        
            'Print #1, vbNewLine & "[sRcv]" & sRcvData;
            
            If InStr(1, sRcvData, "<?xml version") > 0 Then
                 SaveTransDataW = 1
            Else
                SaveTransDataW = -1
            End If
            
            'strAllResult = strAllResult & strEqpCd & "" & sResult & "" & sCommt & "" & Format(Now, "yyyymmddhhmmss") & ""
            
        Next
                    
        SaveTransDataW = 1
    
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
'    cn_Ser.RollbackTrans
    
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
'    strTestCd = mGetP(frmInterface.cboTest.Text, 2, "|")
    pFrDt = Format(frmInterface.dtpStartDt.Value, "yyyymmdd")
    pToDt = Format(frmInterface.dtpStopDt.Value, "yyyymmdd")
'    pFrNo = frmInterface.txtStartNum.Text
'    pToNo = frmInterface.txtStopNum.Text
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'    '-- 전북대병원  r010m.SPCCD
'    SQL = ""
'    SQL = SQL & "SELECT '1', '' AS SN ,'' AS 결과일시, j011m.colldt AS 접수일자, j011m.bcno AS 바코드번호, j010m.bcprtno AS 차트번호" & vbCr
'    SQL = SQL & "       , r010m.WKYMD||r010m.WKGRPCD||r010m.WKNO FLWKNO " & vbCr
'    SQL = SQL & "       , r010m.WKNO AS 접수번호" & vbCr
'    SQL = SQL & "       , j011m.regno AS 내원번호" & vbCr
'    SQL = SQL & "       , j010m.patnm AS 이름" & vbCr
'    SQL = SQL & "       , j010m.age AS 나이" & vbCr
'    SQL = SQL & "       , j010m.sex AS 성별" & vbCr
'    SQL = SQL & "       , j011m.IOGBN" & vbCr
'
'
'    SQL = SQL & " WHERE j011m.colldt BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
'    SQL = SQL & "   AND j011m.bcno = '" & sBarcode & "'" & vbCr


'          SQL = "SELECT '1', '' AS SN ,'' AS 결과일시, a.BSDATE AS 접수일자, a.SAMPLE  AS 바코드번호, a.HOSPNO  AS 차트번호,b.NAME AS 이름,b.SEX AS 성별" & vbCr
          SQL = "SELECT a.BSDATE AS BSDATE, a.SAMPLE  AS SAMPLE, a.HOSPNO  AS HOSPNO,b.NAME AS NAME,b.SEX AS SEX" & vbCr
    SQL = SQL & "  FROM TL_LABOORDER a, TB_IDMAST b" & vbCr
    SQL = SQL & " WHERE a.HOSPNO = b.HOSPNO" & vbCr
    'SQL = SQL & "   AND a.BSDATE BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
    SQL = SQL & "   AND a.SAMPLE = '" & sBarcode & "'" & vbCr
    'SQL = SQL & "   AND a.PUMMOK IN (" & gAllExam & ") " & vbCr
    SQL = SQL & " GROUP BY a.SAMPLE,a.HOSPNO, b.NAME,b.SEX,b.JUMIN,a.BSDATE,a.INOUT  "
                                

    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            'GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
            
            SetText .vasID, "1", .vasID.MaxRows, colCHECKBOX
            SetText .vasID, Trim(RS.Fields("BSDATE")) & "", .vasID.MaxRows, colHOSPDATE
            SetText .vasID, Trim(RS.Fields("SAMPLE")) & "", .vasID.MaxRows, colBARCODE
            SetText .vasID, Trim(RS.Fields("HOSPNO")) & "", .vasID.MaxRows, colCHARTNO
            SetText .vasID, Trim(RS.Fields("NAME")) & "", .vasID.MaxRows, colPNAME
            SetText .vasID, Trim(RS.Fields("SEX")) & "", .vasID.MaxRows, colPSEX
            'SetText .vasID, Trim(RS.Fields("나이")) & "", .vasID.MaxRows, colPAGE
            'SetText .vasID, Trim(RS.Fields("SPCCD")) & "", .vasID.MaxRows, colDISKNO
            
            '-- 화면에 표시
'            For intCol = colState + 1 To .vasID.MaxCols
'                If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                    .vasID.Row = asRow
'                    .vasID.Col = intCol
'                    .vasID.BackColor = vbYellow
'                    Exit For
'                End If
'            Next
    
            RS.MoveNext
        Loop
    
        GetSampleInfoW_SUNGMO = 1
    
    End With
    
'    If GetOrderExamCode <> "" Then
'        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
'        gOrderExam = GetOrderExamCode
'    End If
        
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function



'-- 검사자 정보 가져오기
Function GetSampleInfoW_KNUH(ByVal asRow As Long) As Integer
    Dim sBarcode            As String
    Dim strGubun            As String
    Dim intCol              As Integer
    Dim GetOrderExamCode    As String
    Dim RS1                 As ADODB.Recordset
    Dim strRegDate          As String
    Dim lngRegNo            As Long
    
    
    Dim sParam As String
    Dim sRcvData, sData As String
    Dim varRcvData As Variant
    Dim varTstCode As Variant
    Dim i As Integer
    Dim strTstCD As String
    Dim strItems As String
    Dim intRow As Integer
    
On Error GoTo ErrorTrap

    GetSampleInfoW_KNUH = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'http://병원서버IP/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetMWLBCDJoHoi&Data1=바코드
    
    sParam = "http://"
    sParam = sParam & gKNUH.IP & gKNUH.ORDURL & "&"
    sParam = sParam & "Data1=" & sBarcode
    
' http://kis.kuh.ac.kr:80/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetMWLBCDJoHoi&Data1=32942000940
    
    '==> 서버로 오더조회
    'Print #1, vbNewLine & "[qParam]" & sParam;
    
    sRcvData = OpenURLWithIE2(sParam, frmInterface.Inet1)
    
'    sRcvData = ""
'    sRcvData = sRcvData & "<?xml version=""1.0"" encoding=""euc-kr"" ?>"
'    sRcvData = sRcvData & "<root>"
'    sRcvData = sRcvData & "<IFordcd>"
'    sRcvData = sRcvData & "<ifordcd>L211301210001▦20130121▦02052495▦안종임▦F▦89▦EM▦응급의료센터/0000▦32942000940▦WB, EDTA▦003▦NGAL(Neutrophil Gelatinase Associated Lipocalin)▦L21115▦064618</ifordcd>"
'    sRcvData = sRcvData & "</IFordcd>"
'    sRcvData = sRcvData & "</root>"

    SetRawData "[BC]" & sRcvData
    

'                    Debug.Print sRcvData
    '-- QC
'''    sRcvData = "<?xml version='1.0' encoding='utf-8'?>"
'''    sRcvData = sRcvData & "<root><spcworklist><worklist><acptdt><![CDATA[20120314112224]]></acptdt><bcno><![CDATA[Q24IL0030]]></bcno>"
'''    sRcvData = sRcvData & "<testcd><![CDATA[LIA19601|LIA19604|LIA19616|LIA19617|LIA19606|]]></testcd><testnm><![CDATA[RNP/Sm|RNP(A)|Chromatin|Scl-70|Ro-52 (52 kDa)|]]></testnm>"
'''    sRcvData = sRcvData & "<matrcd><![CDATA[BIO LOW]]></matrcd><matrnm><![CDATA[Bio-plex Low]]></matrnm><levlcd><![CDATA[74]]></levlcd></worklist>"
'''    sRcvData = sRcvData & "</spcworklist></root>"
'''
'''    sRcvData = "<?xml version='1.0' encoding='utf-8'?>"
'''    sRcvData = sRcvData & "<root>"
'''    sRcvData = sRcvData & "<spcworklist>"
'''    sRcvData = sRcvData & "<worklist>"
'''    sRcvData = sRcvData & "<spcacptdt><![CDATA[20120309163857]]></spcacptdt>"
'''    sRcvData = sRcvData & "<acptflag><![CDATA[외래]]></acptflag>"
'''    sRcvData = sRcvData & "<bcno><![CDATA[O24IG2ZL0]]></bcno>"
'''    sRcvData = sRcvData & "<pid><![CDATA[25096972]]></pid>"
'''    sRcvData = sRcvData & "<patnm><![CDATA[문소현]]></patnm>"
'''    sRcvData = sRcvData & "<sexage><![CDATA[F/18]]></sexage>"
'''    sRcvData = sRcvData & "<erprcpflag><![CDATA[N]]></erprcpflag>"
'''    sRcvData = sRcvData & "<workno><![CDATA[20120309I100334]]></workno>"
'''    sRcvData = sRcvData & "<tsectnm><![CDATA[면역부]]></tsectnm>"
'''    sRcvData = sRcvData & "<ifreqcdlist><![CDATA[▦▦▦▦▦▦▦▦▦▦▦]]></ifreqcdlist>"
'''    sRcvData = sRcvData & "<tclscdlist><![CDATA[LIA196▦LIA19601▦LIA19602▦LIA19603▦LIA19604▦LIA19605▦LIA19606▦LIA19608▦LIA19609▦LIA19611▦LIA19614▦]]></tclscdlist>"
'''    sRcvData = sRcvData & "<urinextrvol><![CDATA[ ]]></urinextrvol>"
'''    sRcvData = sRcvData & "<retestyn><![CDATA[N▦N▦N▦N▦N▦N▦N▦N▦N▦N▦N▦]]></retestyn>"
'''    sRcvData = sRcvData & "<rsltstat><![CDATA[LIA196-▦LIA19601-▦LIA19602-▦LIA19603-▦LIA19604-▦LIA19605-▦LIA19606-▦LIA19608-▦LIA19609-▦LIA19611-▦LIA19614-▦]]></rsltstat>"
'''    sRcvData = sRcvData & "</worklist><resultKM error=""no"" type=""status"" clear=""true"" description=""info||정상적으로 처리되었습니다."" updateinstance=""true"" source=""1331617793312""/>"
'''    sRcvData = sRcvData & "</spcworklist></root>"

'L901604290025▦20160429▦02006364▦이숙진▦F▦37▦OBGY▦null▦14136012640▦Vaginal discharge▦168▦Ureaplasma parvum▦L9029407▦120325▩L901604290025▦20160429▦02006364▦이숙진▦F▦37▦OBGY▦null▦14136012640▦Vaginal discharge▦168▦Ureaplasma urealyticum▦L9029406▦120325▩L901604290025▦20160429▦02006364▦이숙진▦F▦37▦OBGY▦null▦14136012640▦Vaginal discharge▦168▦Trichomanas vaginalis▦L9029405▦120325▩L901604290025▦20160429▦02006364▦이숙진▦F▦37▦OBGY▦null▦14136012640▦Vaginal discharge▦168▦STD multiplex Real-time PCR▦L90294▦120325▩L901604290025▦20160429▦02006364▦이숙진▦F▦37▦OBGY▦null▦14136012640▦Vaginal discharge▦168▦Mycoplasma genitalium▦L9029403▦120325▩L901604290025▦20160429▦02006364▦이숙진▦F▦37▦OBGY▦null▦14136012640▦Vaginal discharge▦168▦Neisseria gonorrhoeae▦L9029402▦120325▩L901604290025▦20160429▦02006364▦이숙진▦F▦37▦OBGY▦null▦14136012640▦Vaginal discharge▦168▦Chlamydia trachomatis▦L9029401▦120325▩L901604290025▦20160429▦02006364▦이숙진
'▦F▦37▦OBGY▦null▦14136012640▦Vaginal discharge▦168▦Mycoplasma hominis▦L9029404▦120325


    If InStr(1, sRcvData, "<?xml version") > 0 Then
        ''gwTmp1 = ""
        'varRcvData = Split(sRcvData, "CDATA[")
        
        sRcvData = Mid(sRcvData, InStr(sRcvData, "<ifordcd>") + 9)
        sRcvData = Mid(sRcvData, 1, InStr(sRcvData, "</ifordcd>") - 1)
        varRcvData = Split(sRcvData, "▦")
    End If
    
    If UBound(varRcvData) >= 0 Then
        With frmInterface.vasID
            intRow = .MaxRows
            .Row = intRow
            '.Col = 7
            '.BackColor = vbGreen '&HC6FEFF '&H80C0FF
                                            
            .SetText colCHECKBOX, intRow, "1"
            .SetText colHOSPDATE, intRow, Format(Mid(varRcvData(1), 1, 8), "####-##-##")
            '.SetText colINOUT, intRow, varRcvData(2) & ""
            '.SetText colBARCODE, intRow, varRcvData(8) & ""
            .SetText colCHARTNO, intRow, varRcvData(2) & ""
            .SetText colPNAME, intRow, varRcvData(3) & ""
            .SetText colPSEX, intRow, varRcvData(4) & ""
            .SetText colPAGE, intRow, varRcvData(5) & ""
            '.SetText colER, intRow, varRcvData(7) & ""
'            .SetText colWORKNO, intRow, varRcvData(8) & ""
'            .SetText colPARTNM, intRow, varRcvData(9) & ""
'            .SetText colASSAYNM, intRow, "FOOD"
            '.SetText 13, intRow, varRcvData(11) & ""    'strTstCD
            '.SetText 14, intRow, varRcvData(12) & ""
            '.SetText 15, intRow, varRcvData(13) & ""
            '.SetText 16, intRow, varRcvData(14) & ""
            
            .RowHeight(-1) = 12
            'gRow = intRow
        End With
    End If
        
        
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
    
    frmInterface.vasID.RowHeight(-1) = 12

Exit Function
ErrorTrap:
    GetSampleInfoW_KNUH = -1

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


