Attribute VB_Name = "modDbLibrary"
Option Explicit


Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim VallsID            As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim strDate         As String
    Dim strExamDate     As String
    
    
    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        VallsID = Val(lsID)
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strDate = Format(CDate(.dtpToday.Value), "yyyymmdd")
        strExamDate = Format(CDate(.dtpToday.Value), "yyyy-mm-dd")
        
        '-- Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '장비코드
        SQL = SQL & "   AND EXAMDATE = '" & strDate & "'  " & vbCrLf                                        '검사일
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf     '바코드
        'SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf         'DISK 번호
        'SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "                    'POS 번호
              
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
'                SetRawData "[SQL]" & SQL
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        cn_Ser.BeginTrans
        
        '-- 서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            'sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '결과(수정결과)
            
            '-- 장비결과적용
            'If .optSaveResult(0).Value = True Then
            '    sResult = sResult1
            'Else
                sResult = sResult2
            'End If
            
            If sResult <> "" Then
                              SQL = " Update SLA_LabResult  "
                SQL = SQL & vbCrLf & "   Set Result = '" & sResult & "', "
                SQL = SQL & vbCrLf & "       NormalFlag = '0', "
                SQL = SQL & vbCrLf & "       PanicFlag = '0', "
                SQL = SQL & vbCrLf & "       DeltaFlag = '0', "
                SQL = SQL & vbCrLf & "       TransFlag = '1', "
                SQL = SQL & vbCrLf & "       ResultID  = '', "
                SQL = SQL & vbCrLf & "       ResultDate = '" & Trim(Format(Now, "yyyy-mm-dd")) & "', "
                SQL = SQL & vbCrLf & "       ResultTime = '" & Trim(Format(Time, "HH:MM:SS")) & "' "
                SQL = SQL & vbCrLf & " Where SPECIMENNUM = '" & lsID & "' "
'                SQL = SQL & vbCrLf & "   And OrderCode = '" & strEqpCd & "'"
                SQL = SQL & vbCrLf & "   And LabCode = '" & strEqpCd & "'"
                'SQL = SQL & vbCrLf & "   AND OrderCode IN ('B1010','B1020','CBC5','CBC6','CBC7','CBC8') "
                SQL = SQL & vbCrLf & "   And LabCode = '" & strEqpCd & "'"
                
                SQL = SQL & vbCrLf & "   And transflag < '2' "

                SetRawData "[SQL]" & SQL
                Res = SendQuery(gServer, SQL)

                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow

        If Res = 1 Then
                           SQL = " Update SLA_LabMaster "
            SQL = SQL & vbCrLf & "   Set JStatus = '2' "
            SQL = SQL & vbCrLf & " Where SPECIMENNUM = '" & lsID & "' "
            SQL = SQL & vbCrLf & "   And JStatus < '3' "
'            SQL = SQL & vbCrLf & "   And OrderCode IN (" & gAllExam & ") "
            SQL = SQL & vbCrLf & "   AND OrderCode IN ('B1010','B1020','CBC5','CBC6','CBC7','CBC8','D0002050') "
            'SQL = SQL & vbCrLf & "   And LabCode = '" & strEqpCd & "'"
            SQL = SQL & vbCrLf & "   And RECEIPTDATE = '" & strExamDate & "'"
            
            SetRawData "[SQL]" & SQL
            Res = SendQuery(gServer, SQL)
            
            If Res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
        End If
        
        cn_Ser.CommitTrans
        SaveTransDataW = 1
    
    End With

End Function



Function SaveTransDataR(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow            As Integer
    Dim lsID            As String
    Dim lsPid           As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strTmpCd        As String
    Dim strEqpCd        As String
    Dim strSubEqpCd     As String
    Dim sqlRet          As Integer
    
    SaveTransDataR = -1
    
    lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, colBarcode))
    lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, colPID))
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    With frmInterface
        SQL = ""
        SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,'','',PSEX " & vbCrLf & _
              "  FROM PATRESULT " & vbCrLf & _
              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              "   AND EXAMDATE = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
              "   AND BARCODE = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " '& vbCrLf & _
              "   AND DISKNO = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
              "   AND POSNO = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
        
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
        
        sResult1 = ""
        sResult2 = ""
                
'        cn_Ser.BeginTrans
        
'        '서버로 결과값 저장하기
'        For iRow = 1 To .vasTemp.DataRowCnt
'            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '결과(장비결과)
'            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '결과(수정결과)
'
'            '-- 장비결과 치환
'            sResult1 = Replace(sResult1, "<", "")
'            sResult1 = Replace(sResult1, ">", "")
'
'            If sResult1 <> "" Then
'                SQL = ""
'                SQL = SQL + "UPDATE " & gDBTBL_Parm.RSLTTABLE & vbCrLf      '-- 결과테이블
'                SQL = SQL & "   SET "
'                SQL = SQL & gDBCOLUMN_Parm.RESULT & " = '" & sResult1 & "', " & vbCrLf                                      '결과(장비결과)
'                SQL = SQL & gDBCOLUMN_Parm.RESULT & " = '" & sResult2 & "', " & vbCrLf                                      '결과(수정결과)
'                SQL = SQL & gDBCOLUMN_Parm.MACHCD & " = '" & gEquipCode & "', " & vbCrLf                                    '장비코드
'                SQL = SQL & gDBCOLUMN_Parm.USER & " = '" & gEquipCode & "', " & vbCrLf                                      '결과입력자
'                SQL = SQL & gDBCOLUMN_Parm.RSLTDATE & " = SysDate, " & vbCrLf                                               '결과입력일
'                SQL = SQL & " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & lsID & "' " & vbCrLf                              '바코드번호
'                SQL = SQL & "   AND " & gDBCOLUMN_Parm.TESTCD & " = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' " & vbCrLf   '검사코드
'                SQL = SQL & "   AND " & gDBCOLUMN_Parm.PID & " = '" & lsPid & "' " & vbCrLf                                 '환자번호
'                SQL = SQL & "   AND " & gDBCOLUMN_Parm.STATUS & " < '2' "                                                   '결과상태"
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
            
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = ""
            strSubEqpCd = ""
            strTmpCd = Trim(GetText(.vasTemp, iRow, 2))
            If InStr(strTmpCd, "/") > 0 Then
                strEqpCd = Mid(strTmpCd, 1, InStr(strTmpCd, "/") - 1)
                strSubEqpCd = Mid(strTmpCd, InStr(strTmpCd, "/") + 1)
            Else
                strEqpCd = strTmpCd
                strSubEqpCd = ""
            End If
            'gOrderExam = gOrderExam & "'" & Trim(AdoRs_SQL.Fields("Coda")) & "/" & Trim(AdoRs_SQL.Fields("SubCoda")) & "',"

            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '결과(수정결과)
            '-- 장비결과 치환
'            sResult1 = Replace(sResult1, "<", "")
'            sResult1 = Replace(sResult1, ">", "")
            
            If sResult1 <> "" Then
                SQL = ""
                
'               3. 결과입력
'               AP_INF_Bar_Result  @BCID varchar(20),       '검체번호(바코드)
'                           @InstNo  varchar(3),     '장비코드(Advia120 = '008')
'                           @Coda    varchar(30),        '검사코드
'                           @SubCoda varchar(20),    '검사부코드
'                           @Result  varchar(100)    '검사결과
                cn_Ser.Execute "Exec AP_INF_Bar_Result '" & lsID & "','" & gEquip & "','" & strEqpCd & "','" & strSubEqpCd & "','" & sResult2 & "'", sqlRet
                
                'If sqlRet = 1 Then
                '    lblStatus.Caption = "저장 성공!!"
                'Else
                '    lblStatus.Caption = "저장 실패!!"
                'End If
        
            End If
        Next iRow
            
            
    End With
           
'    cn_Ser.CommitTrans
    SaveTransDataR = 1
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfoW(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim strSex  As String
    Dim strAge  As String
    
    Dim strColPtID
    
    Dim ValBarcode As String
    Dim sExamDate  As String
    
    
    sExamDate = Format(frmInterface.dtpToday, "yyyy-mm-dd")

    
    GetSampleInfoW = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '2 샘플 바코드 번호
    ValBarcode = Val(sBarcode)
    
    If sBarcode = "" Then
        Exit Function
    End If
    
          SQL = "SELECT RECEIPTNO, RECEIPTDATE, PTNO, SPECIMENNUM, SNAME  "
    SQL = SQL & " FROM SLA_LabMaster "
'    SQL = SQL & vbCrLf & " WHERE RECEIPTNO = '" & sBarcode & "' "
    SQL = SQL & vbCrLf & " WHERE SPECIMENNUM = '" & sBarcode & "' "
    SQL = SQL & vbCrLf & "   AND OrderCode IN ('B1010','B1020','CBC5','CBC6','CBC7','CBC8') "
   ' SQL = SQL & vbCrLf & "   AND LABCODE IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND JSTATUS < '3'" & vbLf
    SQL = SQL & vbCrLf & "   AND RECEIPTDATE = '" & sExamDate & "' "
    
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT a.RECEIPTNO AS RECEIPTNO"
    SQL = SQL & ",  a.RECEIPTDATE AS RECEIPTDATE"
    SQL = SQL & ", a.PTNO AS PTNO"
    SQL = SQL & ", a.SPECIMENNUM AS SPECIMENNUM"
    SQL = SQL & ", a.SNAME AS SNAME"
    SQL = SQL & "   FROM SLA_LabMaster a,SLA_LabResult b " & vbCr
    SQL = SQL & " WHERE b.LABCODE IN (" & gAllExam & ") " & vbCr
    SQL = SQL & "   AND a.RECEIPTNO = b.RECEIPTNO " & vbCr
    SQL = SQL & "   AND a.ORDERCODE = b.ORDERCODE " & vbCr
    SQL = SQL & "   and a.RECEIPTDATE = b.RECEIPTDATE" & vbCr
    SQL = SQL & "   AND a.SPECIMENNUM = b.SPECIMENNUM" & vbCr
    SQL = SQL & "   AND a.SPECIMENNUM = '" & sBarcode & "'" & vbCr
    SQL = SQL & "   AND a.RECEIPTDATE = '" & sExamDate & "'" & vbCr
    SQL = SQL & "   AND a.JSTATUS < '3'" & vbCr
    SQL = SQL & "  ORDER BY a.RECEIPTDATE "
    
    SetRawData "[GetSampleInfoW]" & SQL

    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colOrdDate    '1  처방일자
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID        '6  검사번호(=병록번호)
        SetText frmInterface.vasID, Trim(gReadBuf(4)), asRow, colPName      '7  검사자명
        GetSampleInfoW = 1
    Else
        GetSampleInfoW = -1
    End If
    
    frmInterface.vasID.RowHeight(-1) = 15
    
    
End Function

''''-- 수진자 정보 가져오기
'''Function GetSampleInfoW(ByVal asRow As Long) As Integer
'''
'''    Dim sBarcode As String
'''    Dim sSpecNo As String
'''    Dim strAge  As String
'''    Dim sqlRet      As Integer
'''    Dim AdoRs_SQL As ADODB.Recordset
'''    Dim strTestCode As String
'''
'''    GetSampleInfoW = -1
'''
'''    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '2 샘플 바코드 번호
'''
'''    If sBarcode = "" Then
'''        Exit Function
'''    End If
'''
'''    '''2. 대상검사항목조회
'''    '''AP_INF_Bar_Order_Coda   @InstNo varchar(3), '장비코드(Cobas 6000 = '008')
'''    '''                @BCID varchar(20)   '검체번호(바코드)
'''    '''
'''    '''조회내용
'''    '''Coda ,       .SubCoda, Sys_Code,      HCode,     PtName, Serial,       Orderdate,  BCID,     ErYn
'''    '''검사코드 , 검사부코드, 장비코드(무시), 병록번호, 환자명, 원무일련번호, 처방일자,   검체번호, 응급항목여부
'''    '''
'''    '''다른 필드는 무시하고 검사코드, 검사부코드, 검체번호 필드만
'''
'''    Set AdoRs_SQL = New ADODB.Recordset
'''    AdoRs_SQL.CursorLocation = adUseClient
'''    AdoRs_SQL.Open cn_Ser.Execute("Exec AP_INF_Bar_Order_Coda '" & gEquip & "','" & sBarcode & "'", sqlRet)
'''
'''    gOrderExam = ""
'''
'''    If sqlRet = 0 Then
'''        'MsgBox "해당차수의 검사는 완료되었습니다. 차수를 확인하세요.", vbOKOnly + vbExclamation
'''        GetSampleInfoW = -1
'''        Exit Function
'''    Else
'''        With frmInterface.vasID
'''            Do Until AdoRs_SQL.EOF
'''                .SetText colPID, asRow, AdoRs_SQL("HCode") & ""         '환자번호
'''                .SetText colPName, asRow, AdoRs_SQL("PtName") & ""    '환자명
'''
'''                '.SetText colState, colPName, AdoRs_SQL("Coda") & ""     '검사코드
'''                'strTestCode = strTestCode & "'" & Trim(AdoRs_SQL.Fields("Coda")) & "/" & Trim(AdoRs_SQL.Fields("SubCoda")) & "',"
'''
'''                If Trim(AdoRs_SQL.Fields("SubCoda")) & "" <> "" Then
'''                    gOrderExam = gOrderExam & "'" & Trim(AdoRs_SQL.Fields("Coda")) & "/" & Trim(AdoRs_SQL.Fields("SubCoda")) & "',"
'''                Else
'''                    gOrderExam = gOrderExam & "'" & Trim(AdoRs_SQL.Fields("Coda")) & Trim(AdoRs_SQL.Fields("SubCoda")) & "',"
'''                End If
'''                AdoRs_SQL.MoveNext
'''            Loop
'''
'''            GetSampleInfoW = 1
'''        End With
'''    End If
'''
'''    'SetRawData "[TC]" & strTestCode
'''   ' SetRawData "[TC]" & gOrderExam
'''
'''    If gOrderExam <> "" Then
'''        gOrderExam = Mid(gOrderExam, 1, Len(gOrderExam) - 1)
'''    End If
'''
'''End Function

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

Function GetEQPMASTERCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String

    GetEQPMASTERCode = ""
    
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
        GetEQPMASTERCode = Trim(gReadBuf(0))
    End If
    
End Function


Function GetGetEQPMASTERCode_CA1500(argEquipCode As String, argPID As String) As String
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

    GetGetEQPMASTERCode_CA1500 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
'    argPID = "1558200030"
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    Res = GetDBSelectColumn(gServer, SQL)
    GetGetEQPMASTERCode_CA1500 = ""
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
    'GetEQPMASTERCode =
    
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
    
    GetGetEQPMASTERCode_CA1500 = strExamCode
    
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
    
    Dim sExamDate  As String
    
    sExamDate = Format(frmInterface.dtpToday, "yyyy-mm-dd")

    GetOrderExamCode_New = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
          SQL = "SELECT b.LABCODE "
    SQL = SQL & " FROM SLA_LabMaster a, SLA_LABRESULT b "
    SQL = SQL & vbCrLf & " WHERE a.SPECIMENNUM = '" & argPID & "' "
'    SQL = SQL & vbCrLf & "   AND a.OrderCode IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND a.OrderCode IN ('B1010','B1020','CBC5','CBC6','CBC7','CBC8') "
'    SQL = SQL & vbCrLf & "   And b.LabCode IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND a.JSTATUS < '3'"
    SQL = SQL & vbCrLf & "   AND a.SPECIMENNUM = b.SPECIMENNUM "
    SQL = SQL & vbCrLf & "   AND a.ORDERDATE = b.ORDERDATE "
    SQL = SQL & vbCrLf & "   AND a.ORDERCODE = b.ORDERCODE "
    SQL = SQL & vbCrLf & "   AND a.RECEIPTDATE = '" & sExamDate & "' "
    
    
    
    SQL = ""
    SQL = SQL & "SELECT b.LABCODE "
    SQL = SQL & "   FROM SLA_LabMaster a,SLA_LabResult b " & vbCr
    SQL = SQL & " WHERE b.LABCODE IN (" & gAllExam & ") " & vbCr
    SQL = SQL & "   AND a.RECEIPTNO = b.RECEIPTNO " & vbCr
    SQL = SQL & "   AND a.ORDERCODE = b.ORDERCODE " & vbCr
    SQL = SQL & "   and a.RECEIPTDATE = b.RECEIPTDATE" & vbCr
    SQL = SQL & "   AND a.SPECIMENNUM = b.SPECIMENNUM" & vbCr
    SQL = SQL & "   AND a.SPECIMENNUM = '" & argPID & "'" & vbCr
    SQL = SQL & "   AND a.RECEIPTDATE = '" & sExamDate & "'" & vbCr
    SQL = SQL & "   AND a.JSTATUS < '3'" & vbCr
    
    SetRawData "[GetOrderExamCode_New]" & SQL
    
    
    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_New = GetOrderExamCode_New & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
    If GetOrderExamCode_New <> "" Then
        GetOrderExamCode_New = Mid(GetOrderExamCode_New, 1, Len(GetOrderExamCode_New) - 1)
    End If
    
    
    
End Function


Function GetGetEQPMASTERCode_E411(argEquipCode As String, argPID As String, Optional intRow As Long) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim sSpecNo     As String
Dim iRow        As Long
Dim SpecNo      As String
    
    GetGetEQPMASTERCode_E411 = ""
    
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
        GetGetEQPMASTERCode_E411 = ""
        Exit Function
    End If
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEQPMASTERCode =
    
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
    
    GetGetEQPMASTERCode_E411 = Mid(strExamCode, 2)
    
End Function



Function GetGetEQPMASTERCode_Architect(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEQPMASTERCode_Architect = ""
    
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
'        GetGetEQPMASTERCode_Architect = ""
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
    GetGetEQPMASTERCode_Architect = strExamCode
    
End Function

'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetGetEQPMASTERCode_CentaurCP(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEQPMASTERCode_CentaurCP = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 샘플 바코드 번호
    
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

    GetGetEQPMASTERCode_CentaurCP = strExamCode
    
End Function


'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetGetEquipExamCode_XN1000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    Dim strCBC As String
    Dim strDiff As String
    
    GetGetEquipExamCode_XN1000 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 샘플 바코드 번호
    
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
    
    'SetRawData "[GetEquipExamCode_XN1000]" & SQL
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""

    strCBC = ""
    strDiff = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            'NRBC%는 오더를 안준다
'            If Trim(gReadBuf(i)) <> "NRBC%" Then
'                strExamCode = strExamCode & "^^^^" & Trim(gReadBuf(i)) & "\"
'            End If
            
            
            If Trim(gReadBuf(i)) = "WBC" Or Trim(gReadBuf(i)) = "RBC" Or Trim(gReadBuf(i)) = "HGB" Or _
                Trim(gReadBuf(i)) = "HCT" Or Trim(gReadBuf(i)) = "MCV" Or Trim(gReadBuf(i)) = "MCH" Or Trim(gReadBuf(i)) = "MCHC" Or _
                Trim(gReadBuf(i)) = "PLT" Or Trim(gReadBuf(i)) = "RDW-SD" Or Trim(gReadBuf(i)) = "RDW-CV" Or Trim(gReadBuf(i)) = "PDW" Or _
                Trim(gReadBuf(i)) = "MPV" Or Trim(gReadBuf(i)) = "P-LCR" Or Trim(gReadBuf(i)) = "PCT" Or Trim(gReadBuf(i)) = "NRBC#" Or Trim(gReadBuf(i)) = "NRBC%" Then
                
                strCBC = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
                
            End If

            If Trim(gReadBuf(i)) = "NEUT#" Or Trim(gReadBuf(i)) = "LYMPH#" Or Trim(gReadBuf(i)) = "MONO#" Or Trim(gReadBuf(i)) = "EO#" Or Trim(gReadBuf(i)) = "BASO#" Or _
                Trim(gReadBuf(i)) = "NEUT%" Or Trim(gReadBuf(i)) = "LYMPH%" Or Trim(gReadBuf(i)) = "MONO%" Or Trim(gReadBuf(i)) = "EO%" Or Trim(gReadBuf(i)) = "BASO%" Or _
                Trim(gReadBuf(i)) = "IG#" Or Trim(gReadBuf(i)) = "IG%" Then
               
                '-- ^^^^LYMPH#\가 두개인 이유는 ETB 를 장비에서 인식하지 못하기 떄문..(그 자리가 230)
                strDiff = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                
            End If
        Else
            Exit For
        End If
    Next

    strExamCode = strCBC & strDiff
    
    '-- 오더가 없을 경우 CBC만 검사하도록 한다.
    If strExamCode = "" Then
        strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
        strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
    End If
    
    If strExamCode <> "" Then
        strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    End If
    
    GetGetEquipExamCode_XN1000 = strExamCode
    
End Function

'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetGetEQPMASTERCode_Cobas6000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEQPMASTERCode_Cobas6000 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 샘플 바코드 번호
    
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
'    SetRawData "[SQL]" & SQL

    strExamCode = ""

    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i)) & "^"
        Else
            Exit For
        End If
    Next
    
''    '-- E411
''    For Each objResult In mAccInfo.Results
''        strIntBase = objResult.IntNm.IntBase
''        strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
''
''        If strIntBase <> strTemp Then
''            If strItems = "" Then
''                strItems = "^^^" & strIntBase
''            Else
''                strItems = strItems & "\^^^" & strIntBase
''            End If
''            strTemp = strIntBase
''        End If
''    Next
''
''    '-- E601
''    If strIntBase <> strTemp Then
''        If strItems = "" Then
''            strItems = "^^^" & strIntBase & "^" & "1"
''        Else
''            strItems = strItems & "\^^^" & strIntBase & "^" & "1"
''        End If
''        strTemp = strIntBase
''    End If

    If strExamCode <> "" Then
        strExamCode = Mid(strExamCode, 2)
    End If
    GetGetEQPMASTERCode_Cobas6000 = strExamCode
    
    
End Function
Function GetGetEQPMASTERCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEQPMASTERCode_AU480 = ""
    
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
    SQL = SQL & "  FROM EQPMASTER "
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

    GetGetEQPMASTERCode_AU480 = strExamCode
    
End Function


Function GetGetEQPMASTERCode(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String

    GetGetEQPMASTERCode = ""
    
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
    'GetEQPMASTERCode =
    
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
    
    GetGetEQPMASTERCode = strExamCode
    
End Function


