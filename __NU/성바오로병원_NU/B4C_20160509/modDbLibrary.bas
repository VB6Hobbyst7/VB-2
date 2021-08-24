Attribute VB_Name = "modDbLibrary"
Option Explicit


Private Function f_subSet_RefVal(ByVal strORCD As String, ByVal strSubCD As String, Optional ByVal strRSLT As String, Optional ByVal strSex As String, Optional ByVal strAge As String) As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
Dim rs_svr As ADODB.Recordset

On Error GoTo ErrorTrap
    
    strRSLT = Replace(strRSLT, "<", "")
    strRSLT = Replace(strRSLT, ">", "")
    f_subSet_RefVal = " "
    
    f_subSet_RefVal = ""
          SQL = "Select REFHIGH, REFLOW  "
    SQL = SQL & "  From EQPMASTER"
    SQL = SQL & " Where EQUIPNO = '" & gEquip & "' "
    SQL = SQL & "   And EXAMCODE =  '" & strORCD & "'"
'    SQL = SQL & "   And SUBCODE =  '" & strSubCD & "'"
    
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If Res > 0 Then
        If IsNumeric(strRSLT) And IsNumeric(Trim(gReadBuf(0))) And IsNumeric(Trim(gReadBuf(1))) Then
            If Val(strRSLT) > Val(Trim(gReadBuf(0))) Then
                f_subSet_RefVal = "H"
            ElseIf Val(strRSLT) < Val(Trim(gReadBuf(1))) Then
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
    Dim strRegDate  As String
    Dim strOrdNm    As String
    Dim strOrdCd    As String
    Dim strReturn   As String
    
    Dim strID       As String
    Dim sUrl        As String
    Dim sHeader     As String
    Dim sBody       As String
    Dim sSTV        As String
    Dim sRcvData    As String
    Dim strAllResult As String
    Dim sParam      As String
    
On Error GoTo ErrHandle
    
    strAllResult = ""
    
    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        strID = Trim(GetText(.vasID, argSpcRow, colPID))
        
        If lsID = "" Then
            Exit Function
        End If
        
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
'        strChtNum = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        strExamDate = Trim(GetText(.vasID, argSpcRow, colEXAMDATE))
        strSaveSeq = Trim(GetText(.vasID, argSpcRow, colSAVESEQ))
        strRegDate = Trim(GetText(.vasID, argSpcRow, colHOSPDATE))
'        strOrdNm = Trim(GetText(.vasID, argSpcRow, colDOB))

'        Select Case strOrdNm
'            Case "INHALANT":    strOrdCd = gAssayNM.INHALANT_CD
'            Case "FOOD":        strOrdCd = gAssayNM.FOOD_CD
'            Case "ATOPY":       strOrdCd = gAssayNM.ATOPY_CD
'        End Select
        
        
        '-- Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf                            '장비코드
        SQL = SQL & "   AND DISKNO  = '" & strOrdNm & "'" & vbCrLf                          '구분
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCrLf  '검사일
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf                             '바코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq                                        '저장번호
'        SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colBREED)) & "' " & vbCrLf         'DISK 번호(진료검사ID)
'        SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colASSAYNM)) & "' "                    'POS 번호(진료지원ID)
              
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
                
        '-- 서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            strEqpCd = Trim(GetText(.vasTemp, iRow, 1))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))     '결과(IU/ml)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))     '결과(Class)
            'strRefVal = Trim(GetText(.vasTemp, iRow, 5))    '판정
                        
            'strSubCodes = Trim(GetText(.vasTemp, iRow, 14))    '저장용 코드 : ex) 999|888|777

            '-- 장비결과적용
'            If .optSaveResult(0).Value = True Then
'                sResult = sResult1
'            Else
'                sResult = sResult2
'            End If
            
            sResult = sResult1 & " / " & sResult2
                          
            '검사코드,결과값,코멘트,결과보고일자,결과상태 ex) LIA196013201202274LIA1960252012
            strAllResult = strAllResult & strEqpCd & "" & sResult & "" & Format(Now, "yyyymmdd") & "" & "4" & ""
              
              
            'id=1&LabResultList=[{"SerialNo":"B1", "Value":"Value", "Type":1}, {"SerialNo":"B2", "Value":"Value", "Type":2}]
'            If strRefVal = "" Then
'                strRefVal = 1
'            End If
'
'            If lsID <> "" And sResult <> "" And strRegDate <> "" Then
'                sSTV = sSTV & "{"
'                sSTV = sSTV & """SerialNo" & """:""" & strEqpCd & """, "
'                sSTV = sSTV & """Value" & """:""" & sResult1 & """, "
'                sSTV = sSTV & """Type" & """:" & strRefVal & ", "
'                sSTV = sSTV & """Index" & """:" & frmInterface.txtSeq.Text & ", "
'                sSTV = sSTV & """Class" & """:""" & sResult2 & """"
'                sSTV = sSTV & "},"
'            End If
        Next iRow
        
        If strAllResult <> "" Then
                        
            If Mid(lsID, 1, 2) = "QC" Then
                '-- QC검체
                'sParam = "submit_id=TXLQI00101&"
                sParam = "submit_id=TXLQI00102&"
                sParam = sParam & "business_id=lis&"
                sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"      '사용자ID|기관코드
                sParam = sParam & "ifid=" & lsID & "&"                                                  '검체번호(바코드)
                sParam = sParam & "result=" & strAllResult & "&"                                            '결과
                sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                                                 '장비코드
                sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                                    '기관코드
                sParam = sParam & "userid=" & NUAPI.UID & "&"                                             '사용자ID
            Else
                '-- 일반검체
                                   
''        sParam = "submit_id=TRLII00119&"                                        'submit ID
'        sParam = "submit_id=TRLII00101&"                                        'submit ID
                                   
                                   'TXLII00101&
                sParam = "submit_id=TXLII00101&"
                sParam = sParam & "business_id=lis&"
                sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '사용자ID|기관코드
                sParam = sParam & "bcno=" & lsID & "&"                                                  '검체번호(바코드)
                sParam = sParam & "result=" & strAllResult & "&"                                            '결과
                sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                                                '장비코드
                sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                                    '기관코드
                sParam = sParam & "userid=" & NUAPI.UID                                               '사용자ID
            End If
            
            '==> 서버로 결과저장
            'Print #1, vbNewLine & "[sUrl]" & sUrl;
            'Print #1, vbNewLine & "[sParam]" & sParam;
            
            sRcvData = OpenURLWithIE2(NUAPI.APIURL & sParam, frmInterface.Inet1)
            'sRcvData = Inet1.OpenURL(sUrl & sParam)
                        
            'Print #1, vbNewLine & "[sRcv]" & sRcvData;
            
            If InStr(1, sRcvData, "<?xml version") > 0 Then
                 SaveTransDataW = 1
            Else
                SaveTransDataW = -1
            End If
            
            
        Else
            SaveTransDataW = -1
        End If
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
    
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
        SetText frmInterface.vasID, "1", asRow, colChECKBOX
        SetText frmInterface.vasID, sBarcode, asRow, colBARCODE
        'SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colHOSPDATE       '접수일
'        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colCHARTNO        '챠트번호
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID            '등록번호(저장시 필요)
        'SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colDOB          '입/외
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
    
    'SetRawData "[SQL1]" & SQL
    
    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("EXAMCODE")) & "',"
            
            SetText .vasID, "1", asRow, colChECKBOX
            SetText .vasID, sBarcode, asRow, colBARCODE
'            SetText .vasID, Trim(RS.Fields("OSPCHTNUM")), asRow, colCHARTNO         '챠트번호(결과상태 저장시 필요)
            SetText .vasID, Trim(RS.Fields("ResOcmNum")), asRow, colPID             '등록번호(결과     저장시 필요)
            SetText .vasID, Trim(RS.Fields("PbsPatNam")), asRow, colPNAME           '환자명
            
            
            'SetText .vasID, "12345", asRow, colCHARTNO         '챠트번호
            'SetText .vasID, "67890", asRow, colPID            '등록번호(저장시 필요)
            'SetText .vasID, "홍길릴", asRow, colPNAME           '환자명
            
            '-- 화면에 표시
            For intCol = colSTATE + 1 To .vasID.MaxCols
                If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colSTATE, 3) Then
                    .vasID.Row = asRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    '-- 결과저장용 SEQ
                    gArrEquip(intCol - colSTATE, 7) = Trim(RS.Fields("ResOdrSeq")) & "|" & Trim(RS.Fields("ResSeq")) & "|" & Trim(RS.Fields("ResSubSeq"))   '결과저장용 번호's
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

'-- 검사자 정보 가져오기
Function GetSampleInfoW_NTL(ByVal asRow As Long) As Integer
    Dim sBarcode            As String
    Dim strGubun            As String
    Dim intCol              As Integer
    Dim GetOrderExamCode    As String
    Dim RS1                 As ADODB.Recordset
    Dim strRegDate          As String
    Dim lngRegNo            As Long
    
    
    GetSampleInfoW_NTL = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    strRegDate = "20" & Format(Mid(sBarcode, 1, 6), "##-##-##")
    lngRegNo = Val(Mid(sBarcode, 7))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'    If InStr(sBarcode, "-") <= 0 Then
'        Exit Function
'    End If
    
    
    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient    'GetPatientResultList02
    'Set RS = cn_Ser.Execute("Exec Interface_GetPatientResult02 '" & gWKCD & "','" & Format(mGetP(sBarcode, 1, "-"), "####-##-##") & "','" & Val(mGetP(sBarcode, 2, "-")) & "'")
    'Set RS = cn_Ser.Execute("Exec Interface_GetPatientResult02 '" & gWKCD & "','" & Format(Mid(sBarcode, 1, 6), "yyyy-mm-dd") & "','" & Val(Mid(sBarcode, 7)) & "'")
    Set RS = cn_Ser.Execute("Exec Interface_GetPatientResult02 '" & gWKCD & "','" & strRegDate & "','" & lngRegNo & "'")
          
    With frmInterface
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                SetText .vasID, "1", asRow, colChECKBOX
                SetText .vasID, Trim(RS.Fields("LabRegDate")), asRow, colHOSPDATE
'                SetText .vasID, Trim(RS.Fields("PatientChartNo")), asRow, colCHARTNO
                SetText .vasID, Trim(RS.Fields("LabRegNo")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("PatientName")), asRow, colPNAME
'                SetText .vasID, Trim(RS.Fields("CompanyCode")), asRow, colBREED
                SetText .vasID, Trim(RS.Fields("PatientBirthDay")), asRow, colASSAYNM
                SetText .vasID, Trim(RS.Fields("PatientSex")), asRow, colPSEX
                SetText .vasID, Trim(RS.Fields("PatientAge")), asRow, colPAGE
                Select Case Trim(RS.Fields("OrderCode")) & ""
                    Case "62800":   strGubun = "INHALANT"
                    Case "62700":   strGubun = "FOOD"
                    Case "62500":   strGubun = "ATOPY"
                End Select
                
                

                      SQL = " SELECT OrderCode, TestCode, TestSubCode " & vbCrLf
                SQL = SQL & "   FROM LC11_NTL..LabRegResult " & vbCrLf
                SQL = SQL & "  WHERE LABREGDATE = '" & strRegDate & "'" & vbCrLf
                SQL = SQL & "    AND LABREGNO   = " & lngRegNo & vbCrLf
                SQL = SQL & "    AND ORDERCODE  = '" & Trim(RS.Fields("OrderCode")) & "'"
                
                'cn_Ser.CursorLocation = adUseClient
                Set RS1 = cn_Ser.Execute(SQL, , 1)
                If Not RS1.EOF = True And Not RS1.BOF = True Then
                    Do Until RS1.EOF
                        '-- 화면에 표시
                        For intCol = colSTATE + 1 To .vasID.MaxCols
                            If Trim(RS1.Fields("TestSubCode")) = gArrEquip(intCol - colSTATE, 3) And strGubun = gArrEquip(intCol - colSTATE, 7) Then
                                .vasID.Row = asRow
                                .vasID.Col = intCol
                                .vasID.BackColor = vbYellow
                                '-- 결과저장용 SEQ
                                gArrEquip(intCol - colSTATE, 9) = Trim(RS1.Fields("OrderCode")) & "|" & Trim(RS1.Fields("TestCode")) & "|" & Trim(RS1.Fields("TestSubCode"))   '결과저장용 번호's
                                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS1.Fields("TestSubCode")) & "',"
                                Exit For
                            End If
                        Next
                        
                        RS1.MoveNext
                    Loop
                End If
                RS1.Close
                RS.MoveNext
            Loop
        
            GetSampleInfoW_NTL = 1
        
        End If
    End With
        
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
    
    frmInterface.vasID.RowHeight(-1) = 12

End Function



'-- 검사자 정보 가져오기
Function GetSampleInfoW_NU(ByVal asRow As Long) As Integer
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

    GetSampleInfoW_NU = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
          
          
'    If Mid(strBarNo, 1, 1) = "Q" Then
'                 sParam = "submit_id=TRLQI00101&"                                                   'submit ID
'        sParam = sParam & "business_id=lis&"                                                        'business_id
'        sParam = sParam & "ex_interface=" & cUSER_ID & "|" & Trim(txtInstCd.Text) & "&"    '사용자ID|기관코드
'        sParam = sParam & "bcno=" & strBarNo & "&"                                                  '바코드
'        sParam = sParam & "eqmtcd=" & INS_CODE & "&"                                                '장비코드
'        sParam = sParam & "instcd=" & Trim(txtInstCd.Text) & "&"                                    '기관코드
'        sParam = sParam & "userid=" & cUSER_ID                                             '사용자ID
'    Else
        'http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00119&business_id=lis&ex_interface=12345678|012&
        
        sParam = "submit_id=TRLII00101&"                                        'submit ID
        sParam = sParam & "business_id=lis&"                                    'business_id
        sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '사용자ID|기관코드
        sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                          '기관코드
        sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                          '장비코드
        sParam = sParam & "bcno=" & sBarcode                                                        '바코드
'    End If
    
    
'-- 워크리스트
''''        sParam = "submit_id=TRLII00119&"                                        'submit ID
'''        sParam = "submit_id=TRLII00101&"                                        'submit ID
'''        sParam = sParam & "business_id=lis&"                                    'business_id
'''        'sParam = sParam & "refgbn=2&"
'''        sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '사용자ID|기관코드
'''        sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                          '기관코드
'''        sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                           '장비코드
'''        sParam = sParam & "startdd=" & sSch1 & "&"                              '시작작업일자
'''        sParam = sParam & "enddd=" & sSch2 & "&"                                '종료작업일자
'''        'sParam = sParam & "tclscd=" & gAllExam                                  '검사코드리스트
'''        'sParam = sParam & "tclscd=" & strTestCds                                 '검사코드리스트
    
    '==> 서버로 오더조회
    'Print #1, vbNewLine & "[qParam]" & sParam;
    
    sRcvData = OpenURLWithIE2(NUAPI.APIURL & sParam, frmInterface.Inet1)
    
    SetRawData "[BC]" & sRcvData
    
    'Print #1, vbNewLine & "[qRcv]" & sRcvData;

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

    If InStr(1, sRcvData, "<?xml version") > 0 Then
        ''gwTmp1 = ""
        varRcvData = Split(sRcvData, "CDATA[")
    End If
    
    If UBound(varRcvData) >= 0 Then
        For i = 1 To UBound(varRcvData)
            varRcvData(i) = Mid(varRcvData(i), 1, InStr(varRcvData(i), "]") - 1)
        Next
        
        strTstCD = ""
        mOrder.TestCd = ""
        
        If Trim(varRcvData(11) & "") <> "" Then
            varTstCode = Split(varRcvData(11), "▦")
            For i = 0 To UBound(varTstCode) - 1
                strTstCD = strTstCD & "'" & Trim(varTstCode(i)) & "',"
                mOrder.TestCd = mOrder.TestCd & Trim(varTstCode(i)) & "|"
            Next
        End If
                            
        If strTstCD <> "" Then
            strTstCD = Mid(strTstCD, 1, Len(strTstCD) - 1)
        End If
        
'        strItems = ""
'        If Trim(strTstCD) <> "" Then
'            Set adoRS2 = New ADODB.Recordset
'
'                     sqlDoc = "select TESTCD_EQP, TESTNO "
'            sqlDoc = sqlDoc & "  from INTERFACE002 "
'            sqlDoc = sqlDoc & " where EQP_CD = '" & INS_CODE & "'"
'            sqlDoc = sqlDoc & "   and TESTCD in (" & strTstCD & ")"
'            adoRS2.CursorLocation = adUseClient
'            adoRS2.Open sqlDoc, AdoCn_Jet
'            If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
'            Do While Not adoRS2.EOF
'
'                If Trim(adoRS2.Fields("TESTCD_EQP").Value & "") <> "" Then
'                    strChannel = Trim(adoRS2.Fields("TESTCD_EQP").Value & "")
'                    strItems = strItems & "\^^^" & strChannel
'                    varEqpCode = varEqpCode & "|" & strChannel
'                    RecordChk = True
'                End If
'                adoRS2.MoveNext
'            Loop
'
'            Set adoRS2 = Nothing

'        End If
            
'        If strItems <> "" Then
'            strItems = Mid(strItems, 2)
'        End If
        
'        If varEqpCode <> "" Then
'            varEqpCode = Mid(varEqpCode, 2)
'        End If
        
'        If Trim(strItems) = "" Then
'            mOrder.NoOrder = True
'            mOrder.Order = ""
'        Else
'            mOrder.NoOrder = False
'            mOrder.Order = strItems
'        End If
        
        'If RecordChk = True Then
'''    '// 결과가 없고 미통보
'''          SQL = "SELECT a.BSDATE AS BSDATE, a.SAMPLE  AS SAMPLE, a.HOSPNO  AS HOSPNO,b.NAME AS NAME,b.SEX AS SEX" & vbCr
        
        
'spcacptdt 접수일자
'acptflag 입원외래구분
'bcno 검체번호
'PID 등록번호
'patnm 환자명
'sexage 나이성별
'erprcpflag 응급구분
'workno 작업번호
'tsectnm 검사계명
'ifreqcdlist 장비요청코드
'tclscdlist 검사리스트
'urinextrvol 유린값
'retestyn 재검여부
'rsltstat 결과상태
        
        With frmInterface.vasID
            intRow = .MaxRows
            .Row = intRow
            '.Col = 7
            '.BackColor = vbGreen '&HC6FEFF '&H80C0FF
                                            
            .SetText colChECKBOX, intRow, "1"
            .SetText colHOSPDATE, intRow, Format(Mid(varRcvData(1), 1, 8), "####-##-##")
            .SetText colIO, intRow, varRcvData(2) & ""
            .SetText colBARCODE, intRow, varRcvData(3) & ""
            .SetText colPID, intRow, varRcvData(4) & ""
            .SetText colPNAME, intRow, varRcvData(5) & ""
            .SetText colPSEX, intRow, mGetP(varRcvData(6) & "", 1, "/")
            .SetText colPAGE, intRow, mGetP(varRcvData(6) & "", 2, "/")
            .SetText colER, intRow, varRcvData(7) & ""
            .SetText colWORKNO, intRow, varRcvData(8) & ""
            .SetText colPARTNM, intRow, varRcvData(9) & ""
            strTestCds = varRcvData(10) & ""
            strTestCds = Replace(strTestCds, "▦", "")
            .SetText colASSAYNM, intRow, strTestCds
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
    GetSampleInfoW_NU = -1

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
                SetText frmInterface.vasID, "1", asRow, colChECKBOX
                SetText frmInterface.vasID, Mid(mGetP(varResponse(i), 25, vbTab), 1, 8), asRow, colHOSPDATE
                SetText frmInterface.vasID, mGetP(varResponse(i), 0, vbTab), asRow, colBARCODE
                SetText frmInterface.vasID, mGetP(varResponse(i), 7, vbTab), asRow, colPID
                SetText frmInterface.vasID, mGetP(varResponse(i), 26, vbTab), asRow, colPNAME
                
'                Select Case mGetP(varResponse(i), 29, vbTab)
'                    Case "O": SetText frmInterface.vasID, "외래", asRow, colDOB
'                    Case "E": SetText frmInterface.vasID, "응급", asRow, colDOB
'                    Case "I": SetText frmInterface.vasID, "입원", asRow, colDOB
'                End Select
                
                
                For intCol = colSTATE + 1 To .MaxCols
                    If mGetP(varResponse(i), 6, vbTab) = gArrEquip(intCol - colSTATE, 3) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        '-- 결과저장용 SEQ
                        gArrEquip(intCol - colSTATE, 7) = mGetP(varResponse(i), 3, vbTab) & "|" & mGetP(varResponse(i), 4, vbTab) & "|" & mGetP(varResponse(i), 5, vbTab)
                        Exit For
                    End If
                Next intCol
            Next i
        Else
            SetText frmInterface.vasID, "No Order", asRow, colSTATE
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
'        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colDOB
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
            For intCol = colSTATE + 1 To .vasID.MaxCols
                If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colSTATE, 3) Then
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
        
        SpecNo = Trim(GetText(frmInterface.vasID, iRow, colSPECNO))
        
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


