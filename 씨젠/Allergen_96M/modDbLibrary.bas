Attribute VB_Name = "modDbLibrary"
Option Explicit


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
    
On Error GoTo ErrHandle

    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strChtNum = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        strExamDate = Trim(GetText(.vasID, argSpcRow, colEXAMDATE))
        strSaveSeq = Trim(GetText(.vasID, argSpcRow, colSAVESEQ))
        strRegDate = Trim(GetText(.vasID, argSpcRow, colHOSPDATE))
        strOrdNm = Trim(GetText(.vasID, argSpcRow, colINOUT))

        Select Case strOrdNm
            Case "INHALANT":    strOrdCd = gAssayNM.INHALANT_CD
            Case "FOOD":        strOrdCd = gAssayNM.FOOD_CD
            Case "ATOPY":       strOrdCd = gAssayNM.ATOPY_CD
        End Select
        
        
        '-- Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf                            '장비코드
        SQL = SQL & "   AND DISKNO  = '" & strOrdNm & "'" & vbCrLf                          '구분
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCrLf  '검사일
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf                             '바코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq                                        '저장번호
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
        
        '-- 서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))     '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))     '결과(수정결과)
            strRefVal = Trim(GetText(.vasTemp, iRow, 5))    '판정
                        
            strSubCodes = Trim(GetText(.vasTemp, iRow, 14))    '저장용 코드 : ex) 999|888|777

            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If lsID <> "" And strRegDate <> "" And sResult <> "" Then
                '-- 결과업데이트D.  Interface_SetPatientResult02
                cn_Ser.Execute "Exec Interface_SetPatientResult02 '" & strRegDate & "'," & lsPid & ",'" & mGetP(strSubCodes, 1, "|") & "','" & mGetP(strSubCodes, 2, "|") & "','" & mGetP(strSubCodes, 3, "|") & "','" & sResult & "','','',0,0,0,'M010','" & strReturn & "'"
                If Res < 0 Then
                    SaveQuery SQL
                    Exit Function
                Else
                    SaveTransDataW = 1
                End If
                
            End If
        Next iRow
        
        cn_Ser.CommitTrans
        
    
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
    cn_Ser.RollbackTrans
    
End Function



