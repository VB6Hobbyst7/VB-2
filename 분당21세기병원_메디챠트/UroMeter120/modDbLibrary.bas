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
    SQL = SQL & "   And SUBCODE =  '" & strSubCD & "'"
    
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
    Dim VallsID            As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim strSubCD        As String
    Dim strDate         As String
    Dim strRefVal       As String
    Dim strSex As String
    Dim strAge  As String
    Dim strORQN As String
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    
    
    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        VallsID = Val(lsID)
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        'strDate = Format(CDate(.dtpToday.Value), "yyyymmdd")
        strDate = Trim(GetText(.vasID, argSpcRow, colOrdDate))
        
        strYear = Mid(strDate, 1, 4)
        strMonth = Mid(strDate, 5, 2)
        strDay = Mid(strDate, 7, 2)
        
        '-- Local???? ???????? ?????? ????????
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE, SUBCODE " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '????????
        SQL = SQL & "   AND EXAMDATE = '" & strDate & "'  " & vbCrLf                                        '??????
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf     '??????
        'SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf         'DISK ????
        'SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "                    'POS ????
              
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        cn_Ser.BeginTrans
        
        '-- ?????? ?????? ????????
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '????(????????)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '????(????????)
            strSex = Trim(GetText(.vasTemp, iRow, 8))
            strAge = Trim(GetText(.vasTemp, iRow, 10))
            strORQN = Trim(GetText(.vasTemp, iRow, 9))
            strSubCD = Trim(GetText(.vasTemp, iRow, 11)) 'Subcode
            
            '-- ????????????
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If sResult <> "" Then
                '-- H/L ????
                strRefVal = f_subSet_RefVal(strEqpCd, strSubCD, sResult, strSex, strAge)
                
                If strRefVal = "L" Then
                    strRefVal = "2"
                ElseIf strRefVal = "H" Then
                    strRefVal = "1"
                Else
                    strRefVal = "0"
                End If
                
                      SQL = "Update TB_???????? "
                SQL = SQL & "   Set ???????? = '" & sResult & "',"
                SQL = SQL & "       ???????????? = 5, " & vbCr      '1 : ??????, 5 : ????
                SQL = SQL & "       ???????? = '" & strRefVal & "'" & vbCr
                SQL = SQL & " Where ??????   = '" & strYear & "'" & vbCr
                SQL = SQL & "   and ??????   = '" & strMonth & "'" & vbCr
                SQL = SQL & "   and ??????   = '" & strDay & "'" & vbCr
                SQL = SQL & "   and ???????? = '" & lsID & "'" & vbCr
                SQL = SQL & "   And ???????? = '" & strEqpCd & "'" & vbCr
                SQL = SQL & "   And ???????? = '" & strSubCD & "'" & vbCr
            
'                SetRawData "[SQL3]" & SQL
                
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
'?????? ?????? ???????? ????
    Dim iRow            As Integer
    Dim lsID            As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim VallsID         As String
    Dim strDate         As String

    SaveTransDataR = -1
    
    'Local???? ???????? ?????? ????????
    ClearSpread frmInterface.vasTemp
    
    With frmInterface
        lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, 2))
        VallsID = lsID
        lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, 5))
        strDate = Format(CDate(.dtpExamDate.Value), "yyyymmdd")
        
        '-- Local???? ???????? ?????? ????????
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                            '????????
        SQL = SQL & "   AND EXAMDATE = '" & strDate & "'  " & vbCrLf   '??????
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.vasRID, argSpcRow, 2)) & "' " & vbCrLf     '??????
        'SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf         'DISK ????
        'SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "                    'POS ????
                
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
                
        cn_Ser.BeginTrans
        
        '?????? ?????? ????????
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '????(????????)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '????(????????)
            
            '-- ????????????
            If .optSaveResultR(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If sResult <> "" Then
'                If Len(VallsID) > 6 Then
'                    SQL = "Update ONIT..GUMJIN_INTERFACE" & _
'                          "   Set RESULT = '" & sResult & "'," & _
'                          "       ACT_RETURN_DATE = '" & strDate & "'" & _
'                          " Where PER_GUMJIN_DATE = '" & Mid(lsID, 1, 8) & "'" & _
'                          "   And PER_GUM_NUM = " & lsID & "" & _
'                          "   And INTERFACECODE = '" & strEqpCd & "'"
'                Else
                    SQL = "Update onit_out..jun370_resulttb" & _
                          "   Set Result = '" & sResult & "'" & _
                          " Where orderorder = '" & VallsID & "'" & _
                          "   and map2seqno = '" & strEqpCd & "'"
'                End If
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
        
            End If
        Next iRow
            
    End With
           
    cn_Ser.CommitTrans
    SaveTransDataR = 1
    
End Function

'-- ?????? ???? ????????
Function GetSampleInfoW(ByVal asRow As Long) As Integer
    
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim strSex  As String
    Dim strAge  As String
    
    Dim strColPtID
    
    Dim ValBarcode As String
    
    GetSampleInfoW = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '2 ???? ?????? ????
    'ValBarcode = Val(sBarcode)
    
    If sBarcode = "" Then
        Exit Function
    End If
    
                '-- ????????,????????????,??????,????????,????????,????????,????,????,????????,????????
          SQL = "Select DISTINCT a.ORDT,'0' AS ORQN, b.PANM,a.PAID,a.OIFL,'0',b.SEXS,b.AGES,a.NWNO "
    SQL = SQL & "  From LRESULT a, APATINF b"
    SQL = SQL & " Where a.PAID = b.PAID "
    SQL = SQL & "   And a.SPNO =  '" & sBarcode & "'"
    SQL = SQL & "   And a.ORCD in (" & gAllExam & ")"
    SQL = SQL & "   And a.OKFL <> 'Y' "   '-- ????????????
    
    
    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, "1", asRow, colCheckBox
        'SetText frmInterface.vasID, CStr(asRow), asRow, colSeqNo
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colOrdDate
        'SetText frmInterface.vasID, Trim(gReadBuf(3)) & "", asRow, colBarcode
        'SetText vasWorkList, "", intRow, colRack
        'SetText vasWorkList, "", intRow, colPos
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPName
        
        SetText frmInterface.vasID, Trim(gReadBuf(6)), asRow, colSex
        SetText frmInterface.vasID, Trim(gReadBuf(7)), asRow, colAge
        
        GetSampleInfoW = 1
    Else
        GetSampleInfoW = -1
    End If

End Function

Function GetSampleInfoR(ByVal asRow As Long) As Integer
    Dim sBarcode As String
    Dim sSpecNo As String

    GetSampleInfoR = -1
    
    '-- ???????? ????????
    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBarcode))   '???? ?????? ????
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- ???????????? ???????? ????????
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
        '-- ?????? ???????? ?????????? ????
        'strSex = IIf(Mid(Trim(gReadBuf(4)), 7, 1) = "1", "M", "F")
        'SetText frmInterface.vasID, strSex, colSex    '7  ????
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '7  ????
        '-- ?????? ???????? ?????????? ????
        'strAge = Format(Now, "yyyy") - Mid(Trim(gReadBuf(3)), 1, 4)
        'SetText frmInterface.vasID, strAge, asRow, colAge
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colSex    '8  ????
        
        GetSampleInfoR = 1
    Else
    
        GetSampleInfoR = -1
    End If
    
End Function

Function GetEquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'?????????? ???????? ???????? ???????? ???????? ????????
'?? ???? ?????? ?????????? 1?????? ????
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
'?????????? ???????? ???????? ???????? ???????? ????????
'?? ???? ?????? ?????????? 1?????? ????
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
    
    '???????? (R:Routin, E:Stat)
    'strStatFg = IIf(pAccInfo.StatFg = "1", "E", "U")
    strStatFg = "U"
    
    
'    strExamCode = STX & "S2210101" & strStatFg & Space(6) & Space(4) & mOrder.RackNo & mOrder.TubePos & mOrder.BarNo & _
                "B" & Space(15) & strExamCode & ETX
    
    strExamCode = "" & "S2210101" & strStatFg & Space(6) & Space(4) & mResult.RackNo & mResult.TubePos & mResult.BarNo & _
                "B" & Space(15) & strExamCode & ""
    
    GetGetEquipExamCode_CA1500 = strExamCode
    
End Function

'?????????? ???????? ???????? ???????? ???????? ????????
'?? ???? ?????? ?????????? 1?????? ????
Function GetOrderExamCode(argEquipCode As String, argPID As String, Optional argOrdDT As String) As String
    Dim i           As Integer
    Dim sExamCode   As String
    Dim strExamCode As String
    Dim sExamCd     As String
    Dim RS As ADODB.Recordset
    Dim ValargPID   As String
    Dim strORQN     As String
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    
    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    strYear = Mid(argOrdDT, 1, 4)
    strMonth = Mid(argOrdDT, 5, 2)
    strDay = Mid(argOrdDT, 7, 2)
    
    'ValargPID = Val(argPID)
'
'             SQL = "Select Distinct a.ORCD, a.ORQN "
'    SQL = SQL & "  From LRESULT a, APATINF b"
'    SQL = SQL & " Where a.PAID = b.PAID "
'    SQL = SQL & "   And a.SPNO =  '" & argPID & "'"
'    SQL = SQL & "   And a.ORCD in (" & gAllExam & ")"
'    SQL = SQL & "   And a.OKFL <> 'Y' "   '-- ????????????
    
    
          SQL = "Select ????????, ???????? "
    SQL = SQL & "  from tb_???????? " & vbCr
    SQL = SQL & " Where ???????? = '" & argPID & "'" & vbCr
    SQL = SQL & "   And ??????   = '" & strYear & "'" & vbCr
    SQL = SQL & "   And ??????   = '" & strMonth & "'" & vbCr
    SQL = SQL & "   And ??????   = '" & strDay & "'" & vbCr
    SQL = SQL & "   And ???????? > 0 " & vbCr
    SQL = SQL & "   And (???????? is null or ???????? = '') " & vbCr
    SQL = SQL & "   And ????????+???????? in (" & gAllExam & ")"
    
'    SetRawData "[SQL2]" & SQL
    
    Set RS = cn_Ser.Execute(SQL)
    Do Until RS.EOF
        GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields(0)) & Trim(RS.Fields(1)) & "',"
        RS.MoveNext
    Loop
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
    End If
    
    
End Function

Function GetOrderExamCode_New(argEquipCode As String, argPID As String) As String
'?????????? ???????? ???????? ???????? ???????? ????????
'?? ???? ?????? ?????????? 1?????? ????
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
'?????????? ???????? ???????? ???????? ???????? ????????
'?? ???? ?????? ?????????? 1?????? ????
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
    
    '-- ???????? 11?????? ?????????????? ?????? ?????? ??????.
    argPID = Mid(argPID, 1, 10)
    
    If Mid(argPID, 1, 2) = "99" Then
        'strExamCode = Proc_Order_LX_QC(argPID)
        
        'iRow = frmInterface.vasID.DataRowCnt
        iRow = intRow
        
        SpecNo = Trim(GetText(frmInterface.vasID, iRow, colSpecNo))
        
        SQL = "SELECT QC_EXMN_CD "
        SQL = SQL & vbCrLf & " FROM SPSLMQMST "
        SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(SpecNo, 3, 3) & "' "     '//// ???? ????
        SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(SpecNo, 6, 3) & "' "     '//// ?????? ????
        SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(SpecNo, 9, 1) & "' "      '//// ???? ????
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
        '???????????? ???????? ????????
        SQL = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
        Res = GetDBSelectColumn(gServer, SQL)
        sSpecNo = Trim(gReadBuf(0))
        
        '-- ???????? ????????
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
'        MsgBox "?????? ????"
        GetGetEquipExamCode_E411 = ""
        Exit Function
    End If
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    
    '-- ?????? ?????????? ???? ????
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
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 ???? ?????? ????
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- ???????? ????????
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
'        '-- ???????????????? ?????????? ???????? ????
'        GetGetEquipExamCode_Architect = ""
'        Exit Function
'    End If
'
'    '-- ?????? "," ??????
'    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ?????? ?????????? ???? ????
    SQL = "          "
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    
    '-- ???? ?????? ???? ???????? ???????? [ASTM Format >> Architect]
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next
    
    '-- ?????? "\" ??????
    GetGetEquipExamCode_Architect = strExamCode
    
End Function

'-- ?????????????? ?????????? ???????? ???????? ????????
Function GetGetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_AU480 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 ???? ?????? ????
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ?????? ?????????? ???? ????
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            'AU480?? ???? ???????? dilution ?????? ???? '0'????
            strExamCode = strExamCode & "0" & Trim(gReadBuf(i)) & "0"
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_AU480 = strExamCode
    
End Function


'-- ?????????????? ?????????? ???????? ???????? ????????
Function GetGetEquipExamCode_CentaurCP(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_CentaurCP = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 ???? ?????? ????
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ?????? ?????????? ???? ????
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

Function GetGetEquipExamCode(argEquipCode As String, argPID As String) As String
'?????????? ???????? ???????? ???????? ???????? ????????
'?? ???? ?????? ?????????? 1?????? ????
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String

    GetGetEquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '-- ???????? 11?????? ?????????????? ?????? ?????? ??????.
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


