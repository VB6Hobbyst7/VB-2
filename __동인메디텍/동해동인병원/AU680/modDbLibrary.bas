Attribute VB_Name = "modDbLibrary"
Option Explicit


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
        'ClearSpread .vasTemp1
         
        SQL = ""
        SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX " & vbCrLf & _
              "  FROM PAT_RES " & vbCrLf & _
              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              "   AND EXAMDATE = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
              "   AND BARCODE = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
              "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
              
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult1 = ""
        sResult2 = ""
        
        cn_Ser.BeginTrans
        
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '결과(수정결과)
            
            'eGFR 제외
            'If strEqpCd <> "L11050502" Then
                '-- 장비결과 치환
            '    sResult1 = Replace(sResult1, "<", "")
            '    sResult1 = Replace(sResult1, ">", "")
            'End If
            
            If sResult1 <> "" Then
                SQL = ""
                '   3-1. 검사정보 MASTER
                      SQL = "UPDATE JAIN_SCP.SCPRST41 SET "
                SQL = SQL & "       SCP41TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                SQL = SQL & "       SCP41SNDYN  = 'N',"                               '고정값 : 'N'
                SQL = SQL & "       SCP41RSTYN  = 'Y',"                               '고정값 : 'Y'
                SQL = SQL & "       SCP41TSTUID = '" & gUserID & "'"                    '검사자사번
                SQL = SQL & " WHERE SCP41SPMNO2 = '" & lsID & "'"                 '바코드번호
                
                SetRawData "[Rx]" & SQL
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                '   3-2. 검사정보 DETAIL
                         SQL = "UPDATE JAIN_SCP.SCPRST42 SET "
                SQL = SQL & "       SCP42TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                SQL = SQL & "       SCP42RSTCD  = 'N',"                               '결과형식 => 숫자 : 'N', 문자 : 'X', 장문 : 'R'
                SQL = SQL & "       SCP42RESULT = '" & sResult1 & "'"                '결과값
                SQL = SQL & " WHERE SCP42SPMNO2 = '" & lsID & "'"                 '바코드번호
                SQL = SQL & "   AND SCP42SUGACD = '" & strEqpCd & "'"              '수가코드

                SetRawData "[Rx]" & SQL
                
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

End Function


Function SaveTransDataR(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow            As Integer
    Dim lsID            As String
    Dim lsPid           As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String

    SaveTransDataR = -1
    
    lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, colBarcode))
    lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, colPID))
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    
    With frmInterface
        
        SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT " & vbCrLf & _
              "  FROM PAT_RES " & vbCrLf & _
              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              "   AND EXAMDATE = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
              "   AND BARCODE = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              "   AND DISKNO = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
              "   AND POSNO = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
        
        SetRawData "[LSQL]" & SQL
        
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult1 = ""
        sResult2 = ""
        
        cn_Ser.BeginTrans
        
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            '서버로 결과값 저장하기
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '결과(수정결과)
            
            'eGFR 제외
            'If strEqpCd <> "L11050502" Then
                '-- 장비결과 치환
            '    sResult1 = Replace(sResult1, "<", "")
            '    sResult1 = Replace(sResult1, ">", "")
            'End If
            
            If sResult1 <> "" Then
                SQL = ""
                '   3-1. 검사정보 MASTER
                      SQL = "UPDATE JAIN_SCP.SCPRST41 SET "
                SQL = SQL & "       SCP41TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                SQL = SQL & "       SCP41SNDYN  = 'N',"                               '고정값 : 'N'
                SQL = SQL & "       SCP41RSTYN  = 'Y',"                               '고정값 : 'Y'
                SQL = SQL & "       SCP41TSTUID = '" & gUserID & "'"                    '검사자사번
                SQL = SQL & " WHERE SCP41SPMNO2 = '" & lsID & "'"                 '바코드번호
                
                SetRawData "[SQL]" & SQL
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                '   3-2. 검사정보 DETAIL
                         SQL = "UPDATE JAIN_SCP.SCPRST42 SET "
                SQL = SQL & "       SCP42TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                SQL = SQL & "       SCP42RSTCD  = 'N',"                               '결과형식 => 숫자 : 'N', 문자 : 'X', 장문 : 'R'
                SQL = SQL & "       SCP42RESULT = '" & sResult1 & "'"                '결과값
                SQL = SQL & " WHERE SCP42SPMNO2 = '" & lsID & "'"                 '바코드번호
                SQL = SQL & "   AND SCP42SUGACD = '" & strEqpCd & "'"              '수가코드
    
                SetRawData "[Rx]" & SQL
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SaveTransDataR = 1
            
            End If
        Next
    End With
           
Exit Function

    cn_Ser.CommitTrans
    SaveTransDataR = 1
    
End Function

'-- 수진자 정보 가져오기
Function GetSampleInfoW(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim strAGE  As String
    
    GetSampleInfoW = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '2 샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '바코드번호로 환자정보 불러오기
'      -- 테이블 사용
          SQL = "SELECT DiSTINCT b.SCP42IDNOA, a.SCP41NAME, a.SCP41SEX, a.SCP41BIRTH,b.SCP42SUGACD "
    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & sBarcode & "'"
    'SQL = SQL & vbCrLf & "   AND b.SCP42SUGACD in (" & strGumCd & ")"
    SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
    
'
'PCODE 검사파트
'JDATE 검사일
'SPMSID  seq 넘버
'SPMNO 샘플넘버(바코드넘버)
'IDNO    차트번호(기본 7자리 + 타급종번호 1자리, 타급종환자가 아닌경우 ' ' 임)
'IDNAME 환자이름
'KWA 과
'WARD 병동
'Sex 성별
'BIRTHDAY 생일
'SUGACD 검사수가코드
'SUGANM 검사수가명칭
'RESULTYN 결과유무
'SENDYN      결과통보 유무
'SPMTIME 검체일시
    
    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID    '5
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName  '6
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '7
        strAGE = Format(Now, "yyyy") - Mid(Trim(gReadBuf(3)), 1, 4)
        SetText frmInterface.vasID, strAGE, asRow, colAge    '8
        GetSampleInfoW = 1
    Else
        GetSampleInfoW = -1
    End If

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
    
    argPID = Mid(argPID, 1, 10)
    
          SQL = "SELECT DiSTINCT b.SCP42SUGACD "
    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & argPID & "'"
    SQL = SQL & vbCrLf & "   AND (b.SCP42RESULT IS NULL OR b.SCP42RESULT = '') "
    
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
            'eGFR은 오더 안줌
            If Trim(gReadBuf(i)) <> "088" Then
                '                                               dilution
                strExamCode = strExamCode & Trim(gReadBuf(i)) & "0"
            End If
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


