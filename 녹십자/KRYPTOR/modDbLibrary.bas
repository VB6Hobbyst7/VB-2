Attribute VB_Name = "modDbLibrary"
Option Explicit

Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim VallsID            As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim strEqpCd        As String
    Dim strDate         As String
    Dim strPtNo         As String
    Dim strWrkKey       As String
    Dim strWrkNo        As String
    Dim strWrkDte       As String
    Dim strLocalIP      As String
    
    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasWorkList, argSpcRow, colBarcode))
        strWrkDte = Trim(GetText(.vasWorkList, argSpcRow, colOrdDate))
        lsPid = Trim(GetText(.vasWorkList, argSpcRow, colPID))
        strDate = Format(CDate(.dtpToday.Value), "yyyymmdd")
        
        '-- Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,WORKKEY,WORKNO " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '장비코드
        SQL = SQL & "   AND EXAMDATE = '" & strDate & "'  " & vbCrLf                                        '검사일
        SQL = SQL & "   AND BARCODE = '" & lsID & "' "        '바코드
              
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult = ""
        
        strLocalIP = frmInterface.StatusBar1.Panels(2).Text
        
        cn_Ser.BeginTrans
        
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))     '검사코드
            sResult = Trim(GetText(.vasTemp, iRow, 3))      '결과
            strWrkKey = Trim(GetText(.vasTemp, iRow, 8))    'Work Key
            strWrkNo = Trim(GetText(.vasTemp, iRow, 9))     'Work no
            
            
            '-- LABDET Update
'                  SQL = " Select * From MCHRES"
'            SQL = SQL & "  Where REQNO  = '" & lsID & "'"
'            SQL = SQL & "    And ITEMCD = '" & strEqpCd & "' "
           
           '       SQL = " Select * From LABDET"
           ' SQL = SQL & "  Where REQNO  = '" & lsID & "'"
           ' SQL = SQL & "    And ITEMCD = '" & strEqpCd & "' "
            
        '    Set RS = cn_Ser.Execute(SQL, , 1)
            
'            If RS.EOF = True Or RS.BOF = True Then
'                  '-- Insert
'                      SQL = " Insert Into LABDET " & vbCr
'                SQL = SQL & "(REQNO,  ITEMCD, WRKKEY, LABEMP,   WRKDTE, " & vbCr
'                SQL = SQL & " WORKNO, INPDTE, INPTME, SEQNO,    LABRES, " & vbCr
'                SQL = SQL & " MCHCD,  MCHNM,  IPADDR, TRANSCYN, DEL) " & vbCr
'                SQL = SQL & " Values " & vbCr
'                SQL = SQL & "('" & lsID & "', '" & strEqpCd & "', '" & strWrkKey & "', '" & gIFUser & "', '" & strWrkDte & "', " & vbCr
'                SQL = SQL & "'" & strWrkNo & "', to_char(sysdate, 'YYYYMMDD'), to_char(sysdate, 'HH24miss'), 1,'" & sResult & "', " & vbCr
'                SQL = SQL & "'KIM', 'KRYPTOR', '" & strLocalIP & "', 'N', 'N') " & vbCr
'            Else
            '-- Update
                  SQL = " Update LABDET Set " & vbCr
            SQL = SQL & "  LABRES = '" & sResult & "'," & vbCr
            SQL = SQL & "  INPDTE = to_char(sysdate, 'YYYYMMDD'), " & vbCr
            SQL = SQL & "  INPTME = to_char(sysdate, 'HH24mi'), " & vbCr
            SQL = SQL & "  LABEMP = '" & gIFUser & "' " & vbCr
            SQL = SQL & " Where REQNO  = '" & lsID & "'"
            SQL = SQL & "   And ITEMCD = '" & strEqpCd & "'"
                
            Res = SendQuery(gServer, SQL)
            
            If Res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
                
            '-- LABCOM Insert
                  SQL = " Select * From LABCOM"
            SQL = SQL & "  Where REQNO  = '" & lsID & "'"
            SQL = SQL & "    And ITEMCD = '" & strEqpCd & "' "
            
            Set RS = cn_Ser.Execute(SQL, , 1)
    
            If RS.EOF = True Or RS.BOF = True Then
                  '-- Insert
                      SQL = " INSERT INTO LABCOM (REQNO,ITEMCD,LABRES,INPDTE,INPTME,LABYN) "
                SQL = SQL & " VALUES ('" & lsID & "','" & strEqpCd & "','" & sResult & "', to_char(sysdate, 'YYYYMMDD'), to_char(sysdate, 'HH24mi'),'Y')"
            Else
                '-- Delete
                      SQL = " Delete From LABCOM " & vbCr
                SQL = SQL & "  Where REQNO  = '" & lsID & "'"
                SQL = SQL & "    And ITEMCD = '" & strEqpCd & "'"
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                '-- Insert
                      SQL = " INSERT INTO LABCOM (REQNO,ITEMCD,LABRES,INPDTE,INPTME,LABYN) "
                SQL = SQL & " VALUES ('" & lsID & "','" & strEqpCd & "','" & sResult & "', to_char(sysdate, 'YYYYMMDD'), to_char(sysdate, 'HH24mi'),'Y')"
 
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
            End If
              
        Next
    End With
    
    cn_Ser.CommitTrans
    SaveTransDataW = 1

End Function

Function SaveTransDataR(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim VallsID            As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim strEqpCd        As String
    Dim strDate         As String
    Dim strPtNo         As String
    Dim strWrkKey       As String
    Dim strWrkNo        As String
    Dim strWrkDte       As String
    Dim strLocalIP      As String
    Dim strBrcCd        As String
    
    With frmInterface
        SaveTransDataR = -1
        
        lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        strWrkDte = Trim(GetText(.vasRID, argSpcRow, colOrdDate))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))
        strDate = Format(CDate(.dtpToday.Value), "yyyymmdd")
        
        '-- Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        
              'SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,WORKKEY,WORKNO " & vbCrLf
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,WORKKEY,WORKNO,SALETEAM " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '장비코드
'        SQL = SQL & "   AND EXAMDATE = '" & strDate & "'  " & vbCrLf                                        '검사일
        SQL = SQL & "   AND EXAMDATE = '" & strWrkDte & "'  " & vbCrLf                                        '검사일
        SQL = SQL & "   AND BARCODE = '" & lsID & "' "        '바코드

        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult = ""
        
        strLocalIP = frmInterface.StatusBar1.Panels(2).Text
                
        cn_Ser.BeginTrans
                
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))     '검사코드
            sResult = Trim(GetText(.vasTemp, iRow, 3))      '결과
            strWrkKey = Trim(GetText(.vasTemp, iRow, 8))    'Work Key
            strWrkNo = Trim(GetText(.vasTemp, iRow, 9))     'Work no
            strBrcCd = Trim(GetText(.vasTemp, iRow, 10))    '영업소코드
            
            'LABDET Update
                  SQL = " Update LABDET Set " & vbCr
            SQL = SQL & "  LABRES = '" & sResult & "'," & vbCr
            SQL = SQL & "  INPDTE = to_char(sysdate, 'YYYYMMDD'), " & vbCr
            SQL = SQL & "  INPTME = to_char(sysdate, 'HH24mi'), " & vbCr
            SQL = SQL & "  LABEMP = '" & gIFUser & "' " & vbCr
            SQL = SQL & " Where REQNO  = '" & lsID & "'"
            SQL = SQL & "   And ITEMCD = '" & strEqpCd & "'"
                
            Res = SendQuery(gServer, SQL)
            
            If Res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
                            
            '-- LABCOM Select
                  SQL = " Select * From LABCOM"
            SQL = SQL & "  Where REQNO  = '" & lsID & "'"
            SQL = SQL & "    And ITEMCD = '" & strEqpCd & "' "
            
            Set RS = cn_Ser.Execute(SQL, , 1)
            
            If RS.EOF = True Or RS.BOF = True Then
                  '-- LABCOM Insert
                      SQL = " INSERT INTO LABCOM (REQNO,ITEMCD,LABRES,INPDTE,INPTME,LABYN) "
                SQL = SQL & " VALUES ('" & lsID & "','" & strEqpCd & "','" & sResult & "', to_char(sysdate, 'YYYYMMDD'), to_char(sysdate, 'HH24mi'),'Y')"
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
            Else
                '-- LABCOM Delete
                      SQL = " Delete From LABCOM " & vbCr
                SQL = SQL & "  Where REQNO  = '" & lsID & "'"
                SQL = SQL & "    And ITEMCD = '" & strEqpCd & "'"
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                '-- LABCOM Insert
                      SQL = " INSERT INTO LABCOM (REQNO,ITEMCD,LABRES,INPDTE,INPTME,LABYN) "
                SQL = SQL & " VALUES ('" & lsID & "','" & strEqpCd & "','" & sResult & "', to_char(sysdate, 'YYYYMMDD'), to_char(sysdate, 'HH24mi'),'Y')"
 
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
            
            
            '-- SPECIALRPT Insert
            'INSERT INTO SPECIALRPT (reqno, brccd, itemcd, labyn, labempno, docempno, inpdte, inptme, regdte, regtme, imgyn)
            'VALUES ([의뢰번호], [영업소코드], [검사코드], 'Y', [결과입력자], [담당전문의], [입력일자], [입력시간], [입력일자], [입력시간], 'N');
            
                  SQL = " INSERT INTO SPECIALRPT (REQNO, BRCCD, ITEMCD, LABYN, LABEMPNO, DOCEMPNO, INPDTE, INPTME, REGDTE, REGTME, IMGYN)"
            SQL = SQL & " VALUES ('" & lsID & "', '" & strBrcCd & "', '" & strEqpCd & "','Y', '" & gIFUser & "', '', to_char(sysdate, 'YYYYMMDD'), to_char(sysdate, 'HH24mi'), to_char(sysdate, 'YYYYMMDD'), to_char(sysdate, 'HH24mi'), 'N')"

            Res = SendQuery(gServer, SQL)
            
            If Res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
        Next
    End With
    
    cn_Ser.CommitTrans
    SaveTransDataR = 1

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
'                          " Where orderorder = '" & lsID & "'" & _
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
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim strSex  As String
    Dim strAge  As String
    
    Dim strColPtID
    
    Dim ValBarcode As String
    
    GetSampleInfoW = -1
    
    sBarcode = Trim(GetText(frmInterface.vasWorkList, asRow, colBarcode))   '2 샘플 바코드 번호
    'ValBarcode = Val(sBarcode)
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- 바코드번호로 환자정보 불러오기
          SQL = " Select * From MCHORDER "
    SQL = SQL & "  Where REQNO = '" & sBarcode & "'"
    SQL = SQL & "    AND ITEMCD IN (" & gAllExam & ") " & vbLf
    
    Set RS = cn_Ser.Execute(SQL, , 1)

    If RS.EOF = True Or RS.BOF = True Then
        GetSampleInfoW = 0
        Exit Function
    End If
    
    Do Until RS.EOF
        SetText frmInterface.vasWorkList, "0", asRow, colCheckBox
        SetText frmInterface.vasWorkList, Trim(RS.Fields("WORKNO").Value) & "", asRow, colWN
        SetText frmInterface.vasWorkList, Trim(RS.Fields("WRKKEY").Value) & "", asRow, colWK
        SetText frmInterface.vasWorkList, Trim(RS.Fields("WRKDTE").Value) & "", asRow, colOrdDate
        SetText frmInterface.vasWorkList, Trim(RS.Fields("REQNO").Value) & "", asRow, colBarcode
'        SetText frmInterface.vasWorkList, Trim(RS.Fields("BRCNM").Value) & "", asRow, colSale
        SetText frmInterface.vasWorkList, Trim(RS.Fields("BRCCD").Value) & "", asRow, colSale
        SetText frmInterface.vasWorkList, Trim(RS.Fields("CSTNM").Value) & "", asRow, colCST
        SetText frmInterface.vasWorkList, Trim(RS.Fields("SAMPNM").Value) & "", asRow, colSPC
        SetText frmInterface.vasWorkList, Trim(RS.Fields("HOSNO").Value) & "", asRow, colPID
        SetText frmInterface.vasWorkList, Trim(RS.Fields("PATNM").Value) & "", asRow, colPName
        GetSampleInfoW = 1
        RS.MoveNext
    Loop
    
    frmInterface.vasWorkList.RowHeight(-1) = 13
    frmInterface.vasWorkList.Row = 1
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfoS(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim strSex  As String
    Dim strAge  As String
    
    Dim strColPtID
    
    Dim ValBarcode As String
    
    GetSampleInfoS = -1
    
    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBarcode))   '2 샘플 바코드 번호
    'ValBarcode = Val(sBarcode)
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- 바코드번호로 환자정보 불러오기
          SQL = " Select * From MCHORDER "
    SQL = SQL & "  Where REQNO = '" & sBarcode & "'"
    SQL = SQL & "    AND ITEMCD IN (" & gAllExam & ") " & vbLf
    
    Set RS = cn_Ser.Execute(SQL, , 1)

    If RS.EOF = True Or RS.BOF = True Then
        GetSampleInfoS = 0
        Exit Function
    End If
    
    Do Until RS.EOF
        SetText frmInterface.vasRID, "0", asRow, colCheckBox
        SetText frmInterface.vasRID, Trim(RS.Fields("WORKNO").Value) & "", asRow, colWN
        SetText frmInterface.vasRID, Trim(RS.Fields("WRKKEY").Value) & "", asRow, colWK
        SetText frmInterface.vasRID, Trim(RS.Fields("WRKDTE").Value) & "", asRow, colOrdDate
        SetText frmInterface.vasRID, Trim(RS.Fields("REQNO").Value) & "", asRow, colBarcode
'        SetText frmInterface.vasRID, Trim(RS.Fields("BRCNM").Value) & "", asRow, colSale
        SetText frmInterface.vasRID, Trim(RS.Fields("BRCCD").Value) & "", asRow, colSale
        SetText frmInterface.vasRID, Trim(RS.Fields("CSTNM").Value) & "", asRow, colCST
        SetText frmInterface.vasRID, Trim(RS.Fields("SAMPNM").Value) & "", asRow, colSPC
        SetText frmInterface.vasRID, Trim(RS.Fields("HOSNO").Value) & "", asRow, colPID
        SetText frmInterface.vasRID, Trim(RS.Fields("PATNM").Value) & "", asRow, colPName
        GetSampleInfoS = 1
        RS.MoveNext
    Loop
    
    frmInterface.vasRID.RowHeight(-1) = 13
    frmInterface.vasRID.Row = 1
    
End Function

Function GetSampleInfoR(ByVal asRow As Long) As Integer
    Dim sBarcode As String
    Dim sSpecNo As String

    GetSampleInfoR = -1
    
    '-- 환자정보 가져오기
    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBarcode))   '샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- 바코드번호로 환자정보 불러오기
          SQL = "SELECT " & gDBCOLUMN_Parm.PID & "," & gDBCOLUMN_Parm.PNAME & "," & gDBCOLUMN_Parm.PSEX & "," & gDBCOLUMN_Parm.PAGE & vbCrLf
    SQL = SQL & "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
    SQL = SQL & " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
    If gDBCOLUMN_Parm.STATUS <> "" Then
        SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
    End If
    If gDBCOLUMN_Parm.RESULT <> "" Then
        SQL = SQL + "   AND (" & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL)"
    End If
    
    Res = GetDBSelectColumn(gServer, SQL)
    
    If Res = 1 Then
        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName
        '-- 성별이 없을경우 주민번호로 찾기
        'strSex = IIf(Mid(Trim(gReadBuf(4)), 7, 1) = "1", "M", "F")
        'SetText frmInterface.vasID, strSex, colSex    '7  성별
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '7  성별
        '-- 나이가 없을경우 주민번호로 찾기
        'strAge = Format(Now, "yyyy") - Mid(Trim(gReadBuf(3)), 1, 4)
        'SetText frmInterface.vasID, strAge, asRow, colAge
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colSex    '8  나이
        
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

Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String
Dim rs_svr As ADODB.Recordset

    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & " SELECT ITEMCD " & vbLf
    SQL = SQL & "   FROM MCHORDER " & vbLf
    SQL = SQL & "  WHERE REQNO =  '" & argPID & "' " & vbLf
    SQL = SQL & "    AND ITEMCD IN (" & gAllExam & ") " & vbLf
            
    Set rs_svr = cn_Ser.Execute(SQL)
    
    Do Until rs_svr.EOF
        GetOrderExamCode = GetOrderExamCode & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
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
    
    sBarcode = Trim(GetText(frmInterface.vasWorkList, intRow, colBarcode))   '2 샘플 바코드 번호
    
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
            strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i)) & "^^1"
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_CentaurCP = strExamCode
    
End Function


'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetGetEquipExamCode_KRYPTOR(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_KRYPTOR = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasWorkList, intRow, colBarcode))   '2 샘플 바코드 번호
    
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
            strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i)) & "^^1"
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_KRYPTOR = Mid(strExamCode, 2)
    
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


